import psutil
import win32gui
import win32gui
import win32process
import time
from datetime import datetime
import os

def get_active_window_title():
    """Get the title of the active window."""
    window_title = None
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        process_name = process.name()
        window_title = win32gui.GetWindowText(hwnd)
        # Check if the active window is a browser
        if process_name in ["chrome.exe", "firefox.exe", "msedge.exe"]:
            return process_name, window_title
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass
    return None, None

def log_website_titles():
    last_app, last_window_title = None, None

    while True:
        try:
            current_app, current_window_title = get_active_window_title()
            if current_app and current_window_title and (current_app != last_app or current_window_title != last_window_title):
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"[{timestamp}] {current_app}: {current_window_title}")
                last_app, last_window_title = current_app, current_window_title
            time.sleep(1)  # Check every second
        except KeyboardInterrupt:
            print("Logging stopped by user.")
            break

if __name__ == "__main__":
    log_website_titles()