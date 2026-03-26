import ctypes
import subprocess
import threading
import keyboard
from clipboard_manager import register_clipboard_hotkeys
import tkinter as tk

# macros.py
# Requirements: pip install keyboard
# On Windows you may need to run this script as Administrator for global hotkeys or to invoke sleep/hibernate.
# Turn the script into a .exe: 
# 1. pip install pyinstaller
# 2. pyinstaller --onefile --noconsole NAMEofSCRIPT.py (in this case macros.py)
# To make it run on windows startup:
# 1. Press Win + R, type "shell:startup", and hit Enter.
# 2. Place the .exe into that folder.

def show_popup(message, duration=2000):
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)

    width, height = 260, 80

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 2) - (height / 2))

    root.geometry(f"{width}x{height}+{x}+{y}")

    label = tk.Label(
        root,
        text=message,
        font=("Segoe UI", 11),
        bg="#222",
        fg="white"
    )
    label.pack(expand=True, fill="both")

    root.after(duration, root.destroy)
    root.mainloop()

# Sleep the PC function (mapped to F8)
def sleep_pc():
    try:
        # Preferred: call PowrProf.SetSuspendState (may require privileges)
        ctypes.windll.PowrProf.SetSuspendState(False, True, False)
    except Exception:
        # Fallback: use rundll32 (behavior can vary by Windows config)
        try:
            subprocess.run("rundll32.exe powrprof.dll,SetSuspendState 0,1,0", shell=True, check=False)
        except Exception as e:
            print("Failed to put system to sleep:", e)

# Lock the workstation function (mapped to F7)
def lock_workstation():
    try:
        ctypes.windll.user32.LockWorkStation()
    except Exception as e:
        print("Failed to lock workstation:", e)

# Kill the foreground process function (mapped to F9)
def kill_foreground_process():
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return
    
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

    PROCESS_TERMINATE = 0x0001
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid.value)

    if handle:
        kernel32.TerminateProcess(handle, 0)
        kernel32.CloseHandle(handle)

if __name__ == "__main__":
    print("Hotkeys active: F8 = Sleep, F9 = Kill process. Press ESC to quit.")

    show_popup("Macros program active", 2000)

    # Register clipboard hotkeys
    register_clipboard_hotkeys()

    # Other hotkeys
    keyboard.add_hotkey("f7", lambda: (print("F7 pressed -> locking..."), lock_workstation()))
    keyboard.add_hotkey("f8", lambda: (print("F8 pressed -> sleeping..."), sleep_pc()))
    keyboard.add_hotkey("f9", lambda: (print("F9 pressed -> killing foreground process..."), kill_foreground_process()))
    keyboard.wait("shift+esc")
    show_popup("Closing macros program", 1000)