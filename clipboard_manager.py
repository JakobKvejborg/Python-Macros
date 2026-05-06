import threading
import time
import keyboard
import pyperclip
import win32api
import win32clipboard
import win32gui
import ctypes

# clipboard_manager.py

MAX_HISTORY = 20   # how many clipboard entries to remember

clipboard_history: list[str] = []
_history_lock = threading.Lock()

# Tracks which history slot we are currently previewing during a Ctrl+V cycle.
# Reset to -1 when Ctrl is released.
_cycle_index = -1
_cycle_lock  = threading.Lock()

# Set to True just before we push something to the clipboard ourselves so the
# listener does not record our own paste as a new entry.
_ignore_next_change = False
_ignore_lock        = threading.Lock()

WM_CLIPBOARDUPDATE = 0x031D


def _clipboard_has_format(*formats: int) -> bool:
    try:
        win32clipboard.OpenClipboard()
        return any(win32clipboard.IsClipboardFormatAvailable(f) for f in formats)
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def is_image_on_clipboard() -> bool:
    return _clipboard_has_format(win32clipboard.CF_BITMAP, win32clipboard.CF_DIB, win32clipboard.CF_DIBV5)


def is_file_on_clipboard() -> bool:
    return _clipboard_has_format(win32clipboard.CF_HDROP)


def _save_to_history(text: str) -> None:
    if not text:
        return
    with _history_lock:
        # Deduplicate: if this text is already at the front, skip.
        if clipboard_history and clipboard_history[0] == text:
            return
        # Remove any older duplicate so the ring doesn't fill with stale copies.
        try:
            clipboard_history.remove(text)
        except ValueError:
            pass
        clipboard_history.insert(0, text)
        if len(clipboard_history) > MAX_HISTORY:
            clipboard_history.pop()
    print(f"[clipboard] Saved: {text[:60]!r}")

def _on_clipboard_change() -> None:
    global _ignore_next_change

    with _ignore_lock:
        if _ignore_next_change:
            _ignore_next_change = False
            return

    if is_image_on_clipboard() or is_file_on_clipboard():
        return

    try:
        text = pyperclip.paste()
        if text:
            _save_to_history(text)
    except Exception:
        pass


def _wnd_proc(hwnd, msg, wparam, lparam):
    if msg == WM_CLIPBOARDUPDATE:
        _on_clipboard_change()
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


def _clipboard_listener() -> None:
    wc               = win32gui.WNDCLASS()
    wc.lpfnWndProc   = _wnd_proc
    wc.lpszClassName = "ClipboardListenerWindow_CM"
    wc.hInstance     = win32api.GetModuleHandle(None)
    win32gui.RegisterClass(wc)
    hwnd = win32gui.CreateWindow(
        wc.lpszClassName, "", 0,
        0, 0, 0, 0, 0, 0, wc.hInstance, None,
    )
    ctypes.windll.user32.AddClipboardFormatListener(hwnd)
    win32gui.PumpMessages()


# Ctrl+V cycling logic
def _commit_paste() -> None:
    global _ignore_next_change, _cycle_index

    with _history_lock:
        if not clipboard_history:
            _cycle_index = -1
            return
        idx   = max(0, min(_cycle_index, len(clipboard_history) - 1))
        entry = clipboard_history[idx]

    with _ignore_lock:
        _ignore_next_change = True

    pyperclip.copy(entry)
    # Small pause so the clipboard write settles before Ctrl+V.
    time.sleep(0.05)
    keyboard.send("ctrl+v")

    with _cycle_lock:
        _cycle_index = -1

    with _history_lock:
        size = len(clipboard_history)
    print(f"[clipboard] Pasted slot {idx + 1}/{size}: {entry[:60]!r}")


def _wait_for_ctrl_release() -> None:
    done = threading.Event()

    def _hook(e):
        if e.name in ("ctrl", "left ctrl", "right ctrl") and e.event_type == keyboard.KEY_UP:
            done.set()
            return False  # remove hook

    keyboard.hook(_hook)
    done.wait(timeout=30)  # safety: give up after 30 s
    keyboard.unhook(_hook)
    _commit_paste()


# One sentinel so we only spawn the ctrl-release waiter once per cycle.
_ctrl_waiter_active = False
_ctrl_waiter_lock   = threading.Lock()


def handle_paste() -> None:
    global _cycle_index, _ctrl_waiter_active

    # Pass-through for non-text clipboard content.
    if is_image_on_clipboard() or is_file_on_clipboard():
        keyboard.send("ctrl+v")
        return

    with _history_lock:
        if not clipboard_history:
            return

    with _cycle_lock:
        _cycle_index += 1   # advance one step in history (wraps implicitly via clamp)

    with _ctrl_waiter_lock:
        if not _ctrl_waiter_active:
            _ctrl_waiter_active = True

            def _waiter_wrapper():
                global _ctrl_waiter_active
                _wait_for_ctrl_release()
                with _ctrl_waiter_lock:
                    _ctrl_waiter_active = False

            threading.Thread(target=_waiter_wrapper, daemon=True).start()

    with _cycle_lock, _history_lock:
        slot = min(_cycle_index, len(clipboard_history) - 1)
        size = len(clipboard_history)
    print(f"[clipboard] Cycling: slot {slot + 1}/{size} selected (release Ctrl to paste)")


def register_clipboard_hotkeys() -> None:
    threading.Thread(target=_clipboard_listener, daemon=True).start()
    keyboard.add_hotkey("ctrl+v", handle_paste, suppress=True)
    print("[clipboard] Active — Ctrl+C to copy, Ctrl+V (hold & repeat) to cycle history.")