# Imports
import pynput
import keyboard as k
import pyperclip
from subprocess import check_call
import win32clipboard
import time
import re


# Function to keep characters, also found in MiscUtils
def keep_characters(string, allowed_characters):
    # Create a regular expression pattern from the allowed string
    pattern = f"[^{re.escape(allowed_characters)}]"

    # Remove all characters not present in the allowed string
    new_string = re.sub(pattern, "", string)
    return new_string



# Simulates a "ctrl + c"
def ctrl_c():
    k.press('ctrl')
    k.press('c')
    time.sleep(0.1)
    k.release('ctrl')
    k.release('c')

# Simulates a "ctrl + a"
def ctrl_a():
    k.press('ctrl')
    k.press('a')
    time.sleep(0.01)
    k.release('ctrl')
    k.release('a')




# Sets up pynput
#mouse = pynput.mouse.Controller()
#keyboard = pynput.keyboard.Controller()


def add_date():
    k.write('79-81 E')


# Removes non-allowed characters
def remove_extra():
    print("Removing non-allowed characters!")

    # Simulates a "ctrl + c"
    k.press_and_release('enter')
    ctrl_a()
    ctrl_c()

    time.sleep(0.1)

    # Gets clipboards
    win32clipboard.OpenClipboard()
    text = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()

    # Filters text to only contain allowed characters
    text = text.replace(",", ".")
    text = text.replace(":", ".")
    text = keep_characters(string=text, allowed_characters="'1234567890.%")

    # Writes textz
    print(text)
    k.write("'" + text)


# Copies data from pdf to clipboard
def copy_to_clip():
    print("Copying!")

    # Simulates a "ctrl + c"
    ctrl_c()

    # Sleeps
    time.sleep(0.1)

    # Gets clipboards
    win32clipboard.OpenClipboard()
    text = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()

    # Runs cleanup on text
    text = text.replace(",", "")

    print(text)

    # Splits and merges text
    split_text = text.split()
    merged = "'" + "	'".join(split_text)

    time.sleep(0.1)

    # Copies to clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(merged)
    win32clipboard.CloseClipboard()



# Adds hotkeys
k.add_hotkey('/', copy_to_clip)
k.add_hotkey('Scroll_Lock', add_date)
k.add_hotkey('Insert', remove_extra)
# Waits
k.wait()