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

SAVE_INTERVAL = 30        # seconds between auto-saves
RESTORE_DELAY = 6         # seconds after wake before restoring (let Windows settle)
GUID_MONITOR_POWER_ON = "{02731015-4510-4526-99E6-E5A17EBD1AEA}"
HPOWERNOTIFY = ctypes.wintypes.HANDLE()
PBT_POWERSETTINGCHANGE = 0x8013


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

saved_positions = {}  # { proc|index: { x, y, w, h, maximized } }
positions_lock = threading.Lock()

def debug_windows():
    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid).name().lower()
        except:
            return
        placement = win32gui.GetWindowPlacement(hwnd)
        rect = placement[4]
        print(f"{proc} | '{title[:40]}' | placement rect: {rect} | showCmd: {placement[1]}")
    win32gui.EnumWindows(callback, None)

class POWERBROADCAST_SETTING(ctypes.Structure):
    _fields_ = [
        ("PowerSettingGuid", ctypes.c_byte * 16),
        ("PowerSettingIndex", ctypes.wintypes.DWORD),
        ("PowerSettingValue", ctypes.wintypes.DWORD),
    ]

# --- Window helpers ---

def get_proc_name(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        return proc.name().lower()
    except Exception:
        return None

def is_relevant_window(hwnd):
    if not win32gui.IsWindowVisible(hwnd):
        return False
    if win32gui.GetWindowTextLength(hwnd) == 0:
        return False
    proc = get_proc_name(hwnd)
    return proc in TRACKED_APPS

def get_window_placement(hwnd):
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        # placement[1] = showCmd, placement[4] = normal position rect
        maximized = placement[1] == win32con.SW_SHOWMAXIMIZED
        rect = placement[4]  # (left, top, right, bottom) in normal state
        return {
            "x": rect[0],
            "y": rect[1],
            "w": rect[2] - rect[0],
            "h": rect[3] - rect[1],
            "maximized": maximized,
        }
    except Exception:
        return None

def restore_window_placement(hwnd, pos):
    try:
        show_cmd = win32con.SW_SHOWMAXIMIZED if pos["maximized"] else win32con.SW_SHOWNORMAL
        placement = (
            0,                          # flags
            show_cmd,                   # showCmd
            (0, 0),                     # ptMinPosition
            (0, 0),                     # ptMaxPosition
            (                           # rcNormalPosition
                pos["x"],
                pos["y"],
                pos["x"] + pos["w"],
                pos["y"] + pos["h"],
            )
        )
        win32gui.SetWindowPlacement(hwnd, placement)
    except Exception as e:
        print(f"Failed to restore window: {e}")

# --- Save / restore logic ---

def snapshot_positions():
    new_positions = {}
    proc_counters = {}
    def callback(hwnd, _):
        if is_relevant_window(hwnd):
            proc = get_proc_name(hwnd)
            count = proc_counters.get(proc, 0)
            proc_counters[proc] = count + 1
            key = f"{proc}|{count}"
            placement = get_window_placement(hwnd)
            if placement:
                new_positions[key] = placement
    win32gui.EnumWindows(callback, None)
    return new_positions

def save_positions():
    positions = snapshot_positions()
    with positions_lock:
        saved_positions.clear()
        saved_positions.update(positions)
    print(f"[{time.strftime('%H:%M:%S')}] Saved {len(positions)} window positions (in memory).")

def restore_positions():
    print(f"[{time.strftime('%H:%M:%S')}] Restoring window positions...")
    restored = 0
    proc_counters = {}
    def callback(hwnd, _):
        nonlocal restored
        if is_relevant_window(hwnd):
            proc = get_proc_name(hwnd)
            count = proc_counters.get(proc, 0)
            proc_counters[proc] = count + 1
            key = f"{proc}|{count}"
            with positions_lock:
                pos = saved_positions.get(key)
            if pos:
                restore_window_placement(hwnd, pos)
                restored += 1
    win32gui.EnumWindows(callback, None)
    print(f"Restored {restored} windows.")

# --- Auto-save thread ---

def auto_save_loop():
    while True:
        time.sleep(SAVE_INTERVAL)
        save_positions()

# --- Sleep/wake detection via WM_POWERBROADCAST ---

PBT_APMSUSPEND         = 0x0004
PBT_APMRESUMEAUTOMATIC = 0x0012
PBT_APMRESUMESUSPEND   = 0x0007
WM_POWERBROADCAST      = 0x0218

def on_wake():
    print(f"[{time.strftime('%H:%M:%S')}] Wake detected, waiting {RESTORE_DELAY}s for displays to settle...")
    time.sleep(RESTORE_DELAY)
    restore_positions()

def power_listener():
    """Hidden window that receives WM_POWERBROADCAST messages."""
    wc = win32gui.WNDCLASS()
    wc.lpszClassName = "PowerListenerWindow"
    wc.hInstance = ctypes.windll.kernel32.GetModuleHandleW(None)
    wc.lpfnWndProc = wnd_proc
    win32gui.RegisterClass(wc)
    hwnd = win32gui.CreateWindow(
        wc.lpszClassName, "", 0,
        0, 0, 0, 0, 0, 0, wc.hInstance, None
    )
    guid = comtypes.GUID(GUID_MONITOR_POWER_ON)
    ctypes.windll.user32.RegisterPowerSettingNotification(
        hwnd, ctypes.byref(guid), 0
    )
    win32gui.PumpMessages()

def wnd_proc(hwnd, msg, wparam, lparam):
    if msg == WM_POWERBROADCAST:
        if wparam == PBT_APMSUSPEND:
            print(f"[{time.strftime('%H:%M:%S')}] Sleep detected.")
            # Don't save here - Windows may have already disturbed window positions
        elif wparam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
            threading.Thread(target=on_wake, daemon=True).start()
        elif wparam == PBT_POWERSETTINGCHANGE:
            setting = ctypes.cast(lparam, ctypes.POINTER(POWERBROADCAST_SETTING)).contents
            if setting.PowerSettingValue == 0:
                print(f"[{time.strftime('%H:%M:%S')}] Monitor turned off.")
            else:
                print(f"[{time.strftime('%H:%M:%S')}] Monitor turned on — restoring positions.")
                threading.Thread(target=on_wake, daemon=True).start()
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

# --- Entry point ---

def start():
    # Initial snapshot into memory
    save_positions()
    # Auto-save thread
    threading.Thread(target=auto_save_loop, daemon=True).start()
    # Sleep/wake listener (blocks on PumpMessages)
    print(f"Watching {len(TRACKED_APPS)} app(s). Auto-saving every {SAVE_INTERVAL}s.")
    print("Press Ctrl+C to quit.")
    power_listener()

def start_window_watcher():
    save_positions()
    threading.Thread(target=auto_save_loop, daemon=True).start()
    threading.Thread(target=power_listener, daemon=True).start()
    print(f"Window watcher active. Watching {len(TRACKED_APPS)} app(s), auto-saving every {SAVE_INTERVAL}s.")