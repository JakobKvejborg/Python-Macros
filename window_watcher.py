import json
import os
import time
import threading
import ctypes
import ctypes.wintypes
import comtypes
import win32gui
import win32con
import win32process
import psutil

# window_watcher.py

# Tracks window positions for specified apps and restores them after sleep/wake
# or monitor power cycles. Positions are persisted to disk so they survive
# crashes and reboots.


SAVE_INTERVAL   = 30   # seconds between auto-saves
RESTORE_DELAY   = 5    # seconds after wake before restoring (let Windows settle)

# These are the power broadcast messages used for sleep/wake detection
PBT_APMSUSPEND         = 0x0004
PBT_APMRESUMEAUTOMATIC = 0x0012
PBT_APMRESUMESUSPEND   = 0x0007
PBT_POWERSETTINGCHANGE = 0x8013
WM_POWERBROADCAST      = 0x0218


# Persist positions next to this script so they survive restarts.
_SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
POSITIONS_FILE  = os.path.join(_SCRIPT_DIR, "window_positions.json")

GUID_MONITOR_POWER_ON = "{02731015-4510-4526-99E6-E5A17EBD1AEA}"

# Apps to track, by process name (lowercase). Add/remove as needed.
TRACKED_APPS = {
    "code.exe",
    "chrome.exe",
    "firefox.exe",
    "slack.exe",
    "discord.exe",
    "notepad.exe",
    "steam.exe",
    "winword.exe",
    "battle.net.exe",
}

saved_positions: dict = {}
positions_lock  = threading.Lock()

# Guard against duplicate restore calls that can arrive when both
# PBT_APMRESUMESUSPEND *and* the monitor-on POWERBROADCAST_SETTING fire
# within a short window of each other.
_restore_timer: threading.Timer | None = None
_restore_lock   = threading.Lock()

# ---------------------------------------------------------------------------
# Win32 structures
# ---------------------------------------------------------------------------

class POWERBROADCAST_SETTING(ctypes.Structure):
    _fields_ = [
        ("PowerSettingGuid", ctypes.c_byte * 16),
        ("PowerSettingIndex", ctypes.wintypes.DWORD),
        ("PowerSettingValue", ctypes.wintypes.DWORD),
    ]

# ---------------------------------------------------------------------------
# Window helpers
# ---------------------------------------------------------------------------

def get_proc_name(hwnd: int) -> str | None:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(pid).name().lower()
    except Exception:
        return None


def is_relevant_window(hwnd: int) -> bool:
    if not win32gui.IsWindowVisible(hwnd):
        return False
    if win32gui.GetWindowTextLength(hwnd) == 0:
        return False
    proc = get_proc_name(hwnd)
    return proc in TRACKED_APPS


def get_window_placement(hwnd: int) -> dict | None:
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        # placement[1] = showCmd, placement[4] = normal position rect (left, top, right, bottom)
        show_cmd = placement[1]
        rect     = placement[4]
        return {
            "x":        rect[0],
            "y":        rect[1],
            "w":        rect[2] - rect[0],
            "h":        rect[3] - rect[1],
            "maximized": show_cmd == win32con.SW_SHOWMAXIMIZED,
            "minimized": show_cmd == win32con.SW_SHOWMINIMIZED,
        }
    except Exception:
        return None


def restore_window_placement(hwnd: int, pos: dict) -> None:
    try:
        if pos["maximized"]:
            show_cmd = win32con.SW_SHOWMAXIMIZED
        elif pos["minimized"]:
            show_cmd = win32con.SW_SHOWMINIMIZED
        else:
            show_cmd = win32con.SW_SHOWNORMAL

        placement = (
            0,          # flags
            show_cmd,   # showCmd
            (0, 0),     # ptMinPosition
            (0, 0),     # ptMaxPosition
            (           # rcNormalPosition
                pos["x"],
                pos["y"],
                pos["x"] + pos["w"],
                pos["y"] + pos["h"],
            ),
        )
        win32gui.SetWindowPlacement(hwnd, placement)
    except Exception as e:
        print(f"[window_watcher] Failed to restore window: {e}")


def _enumerate_windows() -> dict:
    """Return {key: placement} for all currently visible tracked windows."""
    positions: dict    = {}
    proc_counters: dict = {}

    def callback(hwnd, _):
        if not is_relevant_window(hwnd):
            return
        proc  = get_proc_name(hwnd)
        count = proc_counters.get(proc, 0)
        proc_counters[proc] = count + 1
        key   = f"{proc}|{count}"
        p     = get_window_placement(hwnd)
        if p:
            positions[key] = p

    win32gui.EnumWindows(callback, None)
    return positions


def save_positions() -> None:
    positions = _enumerate_windows()
    with positions_lock:
        saved_positions.clear()
        saved_positions.update(positions)
    # Persist to disk so positions survive a crash or reboot.
    try:
        with open(POSITIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(positions, f, indent=2)
    except Exception as e:
        print(f"[window_watcher] Could not write positions file: {e}")
    print(f"[{time.strftime('%H:%M:%S')}] Saved {len(positions)} window positions.")


def load_positions_from_disk() -> None:
    """Load previously saved positions into memory (called at startup)."""
    if not os.path.exists(POSITIONS_FILE):
        return
    try:
        with open(POSITIONS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        with positions_lock:
            saved_positions.clear()
            saved_positions.update(data)
        print(f"[window_watcher] Loaded {len(data)} saved positions from disk.")
    except Exception as e:
        print(f"[window_watcher] Could not read positions file: {e}")


def restore_positions() -> None:
    print(f"[{time.strftime('%H:%M:%S')}] Restoring window positions...")
    restored     = 0
    proc_counters: dict = {}

    def callback(hwnd, _):
        nonlocal restored
        if not is_relevant_window(hwnd):
            return
        proc  = get_proc_name(hwnd)
        count = proc_counters.get(proc, 0)
        proc_counters[proc] = count + 1
        key   = f"{proc}|{count}"
        with positions_lock:
            pos = saved_positions.get(key)
        if pos:
            restore_window_placement(hwnd, pos)
            restored += 1

    win32gui.EnumWindows(callback, None)
    print(f"[window_watcher] Restored {restored} windows.")


def _auto_save_loop() -> None:
    while True:
        time.sleep(SAVE_INTERVAL)
        save_positions()


def _schedule_restore(delay: float = RESTORE_DELAY) -> None:
    global _restore_timer
    with _restore_lock:
        if _restore_timer is not None:
            _restore_timer.cancel()
        _restore_timer = threading.Timer(delay, _do_restore)
        _restore_timer.daemon = True
        _restore_timer.start()


def _do_restore() -> None:
    global _restore_timer
    print(f"[{time.strftime('%H:%M:%S')}] Wake/monitor-on detected — restoring after {RESTORE_DELAY}s delay...")
    restore_positions()
    with _restore_lock:
        _restore_timer = None


def _wnd_proc(hwnd, msg, wparam, lparam):
    if msg == WM_POWERBROADCAST:
        if wparam == PBT_APMSUSPEND:
            # Save *now*, before Windows has moved anything.
            print(f"[{time.strftime('%H:%M:%S')}] Sleep detected — saving positions.")
            save_positions()

        elif wparam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
            _schedule_restore()

        elif wparam == PBT_POWERSETTINGCHANGE:
            try:
                setting = ctypes.cast(lparam, ctypes.POINTER(POWERBROADCAST_SETTING)).contents
                if setting.PowerSettingValue == 0:
                    print(f"[{time.strftime('%H:%M:%S')}] Monitor turned off.")
                else:
                    print(f"[{time.strftime('%H:%M:%S')}] Monitor turned on.")
                    _schedule_restore()
            except Exception:
                pass

    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


def _power_listener() -> None:
    wc               = win32gui.WNDCLASS()
    wc.lpszClassName = "PowerListenerWindow_WW"
    wc.hInstance     = ctypes.windll.kernel32.GetModuleHandleW(None)
    wc.lpfnWndProc   = _wnd_proc
    win32gui.RegisterClass(wc)

    hwnd = win32gui.CreateWindow(
        wc.lpszClassName, "", 0,
        0, 0, 0, 0, 0, 0, wc.hInstance, None,
    )

    guid = comtypes.GUID(GUID_MONITOR_POWER_ON)
    ctypes.windll.user32.RegisterPowerSettingNotification(
        hwnd, ctypes.byref(guid), 0,
    )
    win32gui.PumpMessages()


# Print info about all visible tracked windows — useful for troubleshooting.
def debug_windows() -> None:
    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc  = psutil.Process(pid).name().lower()
        except Exception:
            return
        placement = win32gui.GetWindowPlacement(hwnd)
        rect      = placement[4]
        print(f"  {proc} | '{title[:40]}' | rect={rect} | showCmd={placement[1]}")
    print("=== Visible windows ===")
    win32gui.EnumWindows(callback, None)
    print("=======================")


def start_window_watcher() -> None:
    load_positions_from_disk()
    save_positions()                                          # fresh snapshot

    threading.Thread(target=_auto_save_loop,  daemon=True).start()
    threading.Thread(target=_power_listener,  daemon=True).start()

    print(
        f"[window_watcher] Active — watching {len(TRACKED_APPS)} app(s), "
        f"auto-saving every {SAVE_INTERVAL}s."
    )
