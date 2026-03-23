import threading
import time
import keyboard
import pyperclip
import win32clipboard
import win32api
import win32gui
import ctypes
import struct

clipboard_history = []
MAX_HISTORY = 10
press_count = 0
paste_timer = None
DELAY = 0.20
FILE_ENTRY = "__FILE_COPY__"
_ignore_next_change = False

WM_CLIPBOARDUPDATE = 0x031D

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

def get_file_paths_from_clipboard():
    try:
        win32clipboard.OpenClipboard()
        if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
            return None
        files = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
        return list(files)
    except Exception:
        return None
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass

def set_files_to_clipboard(file_paths):
    """Rebuilds a CF_HDROP structure from a list of file paths and puts it on the clipboard."""
    global _ignore_next_change
    try:
        # Encode file paths as null-separated, double-null-terminated UTF-16LE
        files_bytes = ('\0'.join(file_paths) + '\0\0').encode('utf-16-le')

        # DROPFILES struct: pFiles(4) + pt.x(4) + pt.y(4) + fNC(4) + fWide(4) = 20 bytes
        drop_files_size = 20
        header = struct.pack('IIIII',
            drop_files_size,  # pFiles: offset to file list
            0,                # pt.x
            0,                # pt.y
            0,                # fNC
            1                 # fWide: 1 = unicode paths
        )
        data = header + files_bytes

        _ignore_next_change = True
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, data)
    except Exception as e:
        print(f"Error restoring file clipboard: {e}")
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
        if isinstance(entry, tuple) and entry[0] == FILE_ENTRY:
            print(f"Saved: [file/folder copy] {entry[1]}")
        else:
            print(f"Saved: {entry[:40]}")

def on_clipboard_change():
    global _ignore_next_change
    if _ignore_next_change:
        _ignore_next_change = False
        return
    try:
        if is_file_on_clipboard():
            file_paths = get_file_paths_from_clipboard()
            if file_paths:
                entry = (FILE_ENTRY, file_paths)
                save_to_history(entry)
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

def is_file_entry(entry):
    return isinstance(entry, tuple) and entry[0] == FILE_ENTRY

def resolve_entry(count_snapshot):
    if not clipboard_history:
        return None, None
    if count_snapshot == 1:
        return 0, clipboard_history[0]
    text_entries = [(i, e) for i, e in enumerate(clipboard_history) if not is_file_entry(e)]
    if not text_entries:
        return None, None
    text_index = min(count_snapshot - 2, len(text_entries) - 1)
    i, entry = text_entries[text_index]
    return i, entry

def do_paste(count_snapshot):
    global press_count, _ignore_next_change
    press_count = 0

    raw_index, entry = resolve_entry(count_snapshot)
    if entry is None:
        return

    if is_file_entry(entry):
        file_paths = entry[1]
        set_files_to_clipboard(file_paths)
        time.sleep(0.1)
        keyboard.send("ctrl+v")
        print(f"Pasted [file/folder] {file_paths} [1/{len(clipboard_history)}]")
    else:
        _ignore_next_change = True
        pyperclip.copy(entry)
        keyboard.send("ctrl+v")
        print(f"Pasted [{raw_index+1}/{len(clipboard_history)}]")

def handle_paste():
    global press_count, paste_timer
    if not clipboard_history:
        return
    press_count += 1

    if paste_timer and paste_timer.is_alive():
        paste_timer.cancel()

    snapshot = press_count
    paste_timer = threading.Timer(DELAY, do_paste, args=[snapshot])
    paste_timer.start()

def register_clipboard_hotkeys():
    listener_thread = threading.Thread(target=clipboard_listener, daemon=True)
    listener_thread.start()
    keyboard.add_hotkey("ctrl+v", handle_paste, suppress=True)
    print("Clipboard cycling active.")
    print("Ctrl+C = copy (auto-detected) | Ctrl+V = smart cycle paste.")