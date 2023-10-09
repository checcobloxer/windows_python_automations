import pyautogui as py
import time

# Get all the gui processes
windows = py.getAllWindows()

# Search for Microsoft Word istance
for window in windows:
    if "Word" in window.title:
        # Maximize the windows
        window.maximize()
        # Show Microsoft Word
        window.activate()
        # Click on insert
        py.click(x=181,y=80, clicks=1, interval=1)

        # Click on expression
        py.click(x=1793,y=118, clicks=1, interval=1)
        break
    else:
        print('Microsoft Word not detected')
        input()