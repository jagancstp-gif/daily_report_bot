daily_report_bot.py
====================
A daily-status-report bot for macOS using Chrome + Numbers
SEQUENCE
--------
    1. Open Chrome in a new tab, load a plain-text weather page.
    2. Select-all + copy the page text (keyboard); parse temperature.
    3. Create a new blank Numbers document (AppleScript).
    4. Write header row + data row directly into cells (AppleScript).
    5. Save natively, then export to date-stamped .xlsx (AppleScript).
    6. Screenshot the finished sheet (PyAutoGUI screenshot -- no
       simulated input involved, just a screen capture).

Uses macOS's native `screencapture` CLI tool instead of
    pyautogui.screenshot(). pyautogui's screenshot function depends on
    pyscreeze -> Pillow, and on some Python versions (e.g. 3.14, which is
    very new) that dependency chain isn't installable yet, raising
    PyAutoGUIException. screencapture ships with macOS itself, so this
    sidesteps the dependency entirely and is arguably more reliable for
    Mac-only automation anyway.
