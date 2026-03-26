import threading
import keyboard
import pyperclip
import win32clipboard
import win32api
import win32gui
import ctypes

clipboard_history = []
MAX_HISTORY = 10
press_count = 0
_ignore_next_change = False
WM_CLIPBOARDUPDATE = 0x031D

def is_image_on_clipboard():
    try:
        win32clipboard.OpenClipboard()
        has_image = (
            win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_BITMAP) or
            win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB)
        )
        return bool(has_image)
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass

def is_file_on_clipboard():
    try:
        win32clipboard.OpenClipboard()
        has_files = win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP)
        return bool(has_files)
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass

def save_to_history(entry):
    if len(clipboard_history) == 0 or clipboard_history[0] != entry:
        clipboard_history.insert(0, entry)
        if len(clipboard_history) > MAX_HISTORY:
            clipboard_history.pop()
        print(f"Saved: {entry[:40]}")

def on_clipboard_change():
    global _ignore_next_change
    if is_image_on_clipboard():
        _ignore_next_change = False
        return
    if _ignore_next_change:
        _ignore_next_change = False
        return
    try:
        if is_file_on_clipboard():
            return
        else:
            text = pyperclip.paste()
            if text:
                save_to_history(text)
    except Exception:
        pass

def clipboard_listener():
    wc = win32gui.WNDCLASS()
    wc.lpfnWndProc = wnd_proc
    wc.lpszClassName = "ClipboardListenerWindow"
    wc.hInstance = win32api.GetModuleHandle(None)
    win32gui.RegisterClass(wc)
    hwnd = win32gui.CreateWindow(
        wc.lpszClassName, "", 0,
        0, 0, 0, 0,
        0, 0, wc.hInstance, None
    )
    ctypes.windll.user32.AddClipboardFormatListener(hwnd)
    win32gui.PumpMessages()

def wnd_proc(hwnd, msg, wparam, lparam):
    if msg == WM_CLIPBOARDUPDATE:
        on_clipboard_change()
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

def do_paste():
    global press_count, _ignore_next_change
    if not clipboard_history:
        press_count = 0
        return
    index = min(press_count - 1, len(clipboard_history) - 1)
    entry = clipboard_history[index]
    _ignore_next_change = True
    pyperclip.copy(entry)
    keyboard.send("ctrl+v")
    print(f"Pasted [{index+1}/{len(clipboard_history)}]")
    press_count = 0

def wait_for_ctrl_release():
    done = threading.Event()
    def on_ctrl_up(e):
        if e.name == "ctrl" and e.event_type == keyboard.KEY_UP:
            done.set()
            return False  # unhook
    keyboard.hook(on_ctrl_up)
    done.wait()
    do_paste()

def handle_paste():
    global press_count
    if is_image_on_clipboard() or is_file_on_clipboard():
        keyboard.send("ctrl+v")
        return
    if not clipboard_history:
        return
    press_count += 1
    # Only spawn the release-waiter on the first V press
    if press_count == 1:
        threading.Thread(target=wait_for_ctrl_release, daemon=True).start()

def register_clipboard_hotkeys():
    listener_thread = threading.Thread(target=clipboard_listener, daemon=True)
    listener_thread.start()
    keyboard.add_hotkey("ctrl+v", handle_paste, suppress=True)
    print("Clipboard cycling active.")
    print("Ctrl+C = copy (auto-detected) | Ctrl+V = smart cycle paste.")