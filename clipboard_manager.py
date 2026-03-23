import threading
import time
import keyboard
import pyperclip


clipboard_history = []
MAX_HISTORY = 10 # The number of clipboard entries to remember
press_count = 0 # The number of times the paste hotkey has been pressed, used for debugging
paste_timer = None
DELAY = 0.20 # The delay before pasting the clipboard content, lower = faster paste feel

# Function that handles copying text to the clipboard
def handle_copy():
    global clipboard_history

    time.sleep(0.1)
    text = pyperclip.paste()

    if not text:
        return
    
    if len(clipboard_history) == 0 or clipboard_history[0] != text:
        clipboard_history.insert(0, text)

        # Delete the oldest entry if we exceed the max history
        if len(clipboard_history) > MAX_HISTORY:
            clipboard_history.pop()

        print(f"Saved: {text[:40]}") # Print the first 40 characters of the copied text for debugging

# Function that handles pasting text from the clipboard history
def do_paste():
    global press_count

    if not clipboard_history:
        return
    
    index = min(press_count -1, len(clipboard_history) - 1)
    text = clipboard_history[index]

    pyperclip.copy(text)
    keyboard.send("ctrl+v")

    print(f"Pasted [{index+1}/{len(clipboard_history)}]")

    # Reset after paste
    press_count = 0
          
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

def register_clipboard_hotkeys():
    keyboard.add_hotkey("ctrl+c", handle_copy, suppress=False)
    keyboard.add_hotkey("ctrl+v", handle_paste, suppress=True)

    print("Clipboard cycling active.")
    print("Ctrl+C = copy | Ctrl+V = smart cycle paste.")
