import ctypes
import subprocess
import threading
import sys
import time
import keyboard
 
# macros.py 
# # Requirements: pip install keyboard 
# # On Windows you may need to run this script as Administrator for global hotkeys or to invoke sleep/hibernate. 
# # Turn the script into a .exe: 
# # 1. pip install pyinstaller 
# # 2. pyinstaller --onefile --noconsole NAMEofSCRIPT.py (in this case macros.py) 
# # To make it run on windows startup: 
# # 1. Press Win + R, type "shell:startup", and hit Enter. 
# # 2. Place the .exe into that folder.

try:
    from clipboard_manager import register_clipboard_hotkeys
    from window_watcher    import start_window_watcher
except ImportError as exc:
    print(f"[macros] Import error: {exc}")
    sys.exit(1)
 
def show_popup(message: str, duration: int = 2000) -> None:
    def _run():
        try:
            import tkinter as tk
            root = tk.Tk()
            root.overrideredirect(True)
            root.attributes("-topmost", True)
            root.attributes("-alpha", 0.93)
 
            width, height  = 280, 64
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            # Bottom-right with a small margin (mimics Windows toasts).
            x  = sw - width  - 16
            y  = sh - height - 56   # above the taskbar
            root.geometry(f"{width}x{height}+{x}+{y}")
 
            tk.Label(
                root,
                text=message,
                font=("Segoe UI", 10),
                bg="#1e1e1e",
                fg="#e8e8e8",
                padx=14,
                pady=10,
            ).pack(expand=True, fill="both")
 
            root.after(duration, root.destroy)
            root.mainloop()
        except Exception as e:
            print(f"[macros] Popup error: {e}")
 
    threading.Thread(target=_run, daemon=True).start()
 

def sleep_pc() -> None:
    show_popup("💤  Sleeping…", 1500)
    time.sleep(0.4)   # give the popup a moment to appear before display dies
    try:
        # Preferred: direct WinAPI call.  Args: hibernate=False, forceCritical=True, disableWakeEvent=False
        ctypes.windll.PowrProf.SetSuspendState(False, True, False)
    except Exception:
        try:
            subprocess.run(
                "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",
                shell=False,
                check=False,
            )
        except Exception as e:
            print(f"[macros] Sleep failed: {e}")
 
def lock_workstation() -> None:
    """Lock the current Windows session."""
    print("[macros] Locking workstation...")
    show_popup("🔒  Locking…", 1200)
    time.sleep(0.2)
    try:
        ctypes.windll.user32.LockWorkStation()
    except Exception as e:
        print(f"[macros] Lock failed: {e}")
 
def kill_foreground_process() -> None:
    """Terminate whichever process owns the current foreground window."""
    user32   = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
 
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        print("[macros] No foreground window found.")
        return
 
    pid = ctypes.c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        print("[macros] Could not get PID of foreground window.")
        return
 
    PROCESS_TERMINATE = 0x0001
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid.value)
    if handle:
        kernel32.TerminateProcess(handle, 0)
        kernel32.CloseHandle(handle)
    else:
        print(f"[macros] Could not open process {pid.value} — try running as Administrator.")
 
def _on_f7():
    print("[macros] F7 → lock")
    lock_workstation()
 
def _on_f8():
    print("[macros] F8 → sleep")
    sleep_pc()
 
def _on_f9():
    print("[macros] F9 → kill foreground")
    kill_foreground_process()
 
if __name__ == "__main__":
    print("=" * 50)
    print("  macros.py starting up")
    print("  F7 = Lock | F8 = Sleep | F9 = Kill process")
    print("  Shift+Esc = Quit")
    print("=" * 50)
 
    show_popup("⚡  Macros active", 2000)
 
    # Start sub-systems
    register_clipboard_hotkeys()
    start_window_watcher()
 
    # Register hotkeys
    keyboard.add_hotkey("f7",         _on_f7, suppress=False)
    keyboard.add_hotkey("f8",         _on_f8, suppress=False)
    keyboard.add_hotkey("f9",         _on_f9, suppress=False)
 
    # Block until the quit combo is pressed.
    keyboard.wait("shift+esc")
 
    print("[macros] Shutting down...")
    show_popup("👋  Macros closed", 1000)
    time.sleep(1.1)   # let the popup render before the process exits
    sys.exit(0)
