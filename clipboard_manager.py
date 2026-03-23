import threading
import time
import keyboard
import pyperclip
import win32clipboard

clipboard_history = []
MAX_HISTORY = 10 # The number of clipboard entries to remember
press_count = 0 # The number of times the paste hotkey has been pressed, used for debugging
paste_timer = None
DELAY = 0.20 # The delay before pasting the clipboard content, lower = faster paste feel

FILE_ENTRY = "__FILE_COPY__"  # Sentinel value stored in history when files/folders are copied

# Function to check if the clipboard contains a file/folder
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

# Function that handles copying text to the clipboard
def handle_copy():
    global clipboard_history
    time.sleep(0.1)

    # Check for file/folder copy first
    if is_file_on_clipboard():
        if len(clipboard_history) == 0 or clipboard_history[0] != FILE_ENTRY:
            clipboard_history.insert(0, FILE_ENTRY)
            if len(clipboard_history) > MAX_HISTORY:
                clipboard_history.pop()
            print("Saved: [file/folder copy]")
        return

    text = pyperclip.paste()
    if not text:
        return
    
    if len(clipboard_history) == 0 or clipboard_history[0] != text:
        clipboard_history.insert(0, text)
        # Delete the oldest entry if we exceed the max history
        if len(clipboard_history) > MAX_HISTORY:
            clipboard_history.pop()
        print(f"Saved: {text[:40]}") # Print the first 40 characters of the copied text for debugging

# Actually executes the paste operation
def do_paste():
    global press_count
    if not clipboard_history:
        return
    
    index = min(press_count - 1, len(clipboard_history) - 1)
    entry = clipboard_history[index]

    if entry == FILE_ENTRY:
        # Files are already on the system clipboard — just trigger the native paste
        keyboard.send("ctrl+v")
        print(f"Pasted [file/folder] [{index+1}/{len(clipboard_history)}]")
    else:
        pyperclip.copy(entry)
        keyboard.send("ctrl+v")
        print(f"Pasted [{index+1}/{len(clipboard_history)}]")

    # Reset after paste
    press_count = 0

# Function to handle the timing/cycling
def handle_paste():
    global press_count, paste_timer
    if not clipboard_history:
        return
    press_count += 1
    
    # Cancel previous timer
    if paste_timer and paste_timer.is_alive():
        paste_timer.cancel()
    paste_timer = threading.Timer(DELAY, do_paste)
    paste_timer.start()

# Function to register global hotkeys, used in macros.py
def register_clipboard_hotkeys():
    keyboard.add_hotkey("ctrl+c", handle_copy, suppress=False)
    keyboard.add_hotkey("ctrl+v", handle_paste, suppress=True)
    print("Clipboard cycling active.")
    print("Ctrl+C = copy | Ctrl+V = smart cycle paste.")