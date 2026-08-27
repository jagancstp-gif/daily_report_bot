#!/usr/bin/env python3
"""
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
"""

import os
import re
import time
import subprocess
from datetime import datetime

import pyautogui
import pyperclip

# --------------------------------------------------------------------------
# SAFETY / GLOBAL PYAUTOGUI SETTINGS
# --------------------------------------------------------------------------
pyautogui.FAILSAFE = True   # slam mouse into a screen corner to abort
pyautogui.PAUSE = 0.4


# --------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------
OUTPUT_DIR = os.path.expanduser("~/Desktop/daily_reports")
WEATHER_URL = "https://wttr.in/~Siruseri,+India?format=3" # plain-text weather line
COMMENT_TEXT = "Good for outdoor activities"

APP_LAUNCH_WAIT = 3.0
PAGE_LOAD_WAIT = 4.0


# --------------------------------------------------------------------------
# APPLESCRIPT HELPER
# --------------------------------------------------------------------------
def run_applescript(script: str) -> str:
    """
    Runs an AppleScript snippet via osascript and returns stdout.
    Raises with the stderr message if the script fails, since a silent
    AppleScript failure is exactly what caused confusing behaviour before.
    """
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript failed:\n{script}\n--- stderr ---\n{result.stderr.strip()}")
    return result.stdout.strip()


def activate_app(app_name: str, wait_after: float = APP_LAUNCH_WAIT):
    run_applescript(f'tell application "{app_name}" to activate')
    time.sleep(wait_after)


def escape_for_applescript(text: str) -> str:
    """Escapes double quotes and backslashes so they're safe inside an AS string literal."""
    return text.replace("\\", "\\\\").replace('"', '\\"')


# --------------------------------------------------------------------------
# STEP 1 - FETCH DATA FROM THE WEB (Chrome, new tab) -- keyboard only,
#          this is the one step PyAutoGUI genuinely has to do the work.
# --------------------------------------------------------------------------
def fetch_weather_via_browser() -> str:
    print("Step 1: Opening Chrome in a new tab and fetching weather data...")
    activate_app("Google Chrome")

    pyautogui.hotkey("command", "t")     # force a brand-new tab
    time.sleep(0.8)

    pyautogui.hotkey("command", "l")     # focus address bar
    time.sleep(0.5)
    pyautogui.typewrite(WEATHER_URL, interval=0.02)
    pyautogui.press("enter")

    print(f"[wait] {PAGE_LOAD_WAIT:.1f}s -- letting the weather page load")
    time.sleep(PAGE_LOAD_WAIT)

    pyautogui.press("tab")               # move focus off the omnibox
    time.sleep(0.3)
    pyautogui.hotkey("command", "a")
    time.sleep(0.3)
    pyautogui.hotkey("command", "c")
    time.sleep(0.5)

    raw_text = pyperclip.paste().strip()
    print(f"[debug] Clipboard captured: {raw_text!r}")

    match = re.search(r"[+-]?\d+°C", raw_text)
    temperature = match.group(0) if match else "N/A"
    city_match = re.match(r"([A-Za-z ]+):", raw_text)
    city = city_match.group(1).strip() if city_match else "Unknown location"

    fetched_data = f"{city}: {temperature}"
    print(f"Step 1 complete -> fetched_data = {fetched_data}")
    return fetched_data


# --------------------------------------------------------------------------
# STEP 2 - BUILD THE REPORT ROW
# --------------------------------------------------------------------------
def build_report_row(fetched_data: str):
    now = datetime.now()
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
    date_str = now.strftime("%Y-%m-%d")

    xlsx_filename = f"daily_report_{date_str}.xlsx"
    screenshot_filename = f"daily_report_{date_str}.png"

    return timestamp_str, fetched_data, COMMENT_TEXT, xlsx_filename, screenshot_filename


# --------------------------------------------------------------------------
# STEP 3 - NUMBERS: create doc + write cells DIRECTLY via AppleScript
#          (no simulated typing at all -- this is the actual fix)
# --------------------------------------------------------------------------
def open_numbers_new_sheet():
    print("Step 3: Creating a new blank Numbers document...")
    activate_app("Numbers")
    run_applescript('tell application "Numbers" to make new document')
    time.sleep(APP_LAUNCH_WAIT)


def set_cell(cell_ref: str, value: str):
    """
    Sets a single cell's value directly through Numbers' scripting
    dictionary -- e.g. set_cell("A1", "Timestamp"). No keystrokes
    involved, so there's nothing for an autocomplete popup or a menu
    shortcut to intercept.
    """
    safe_value = escape_for_applescript(value)
    script = f'''
    tell application "Numbers"
        tell table 1 of active sheet of front document
            set value of cell "{cell_ref}" to "{safe_value}"
        end tell
    end tell
    '''
    run_applescript(script)


def write_report_into_sheet(timestamp_str: str, fetched_data: str, comment: str):
    print("Step 3: Writing header + data row directly into cells...")
    # Header row
    set_cell("A1", "Timestamp")
    set_cell("B1", "Data")
    set_cell("C1", "Comment")
    # Data row
    set_cell("A2", timestamp_str)
    set_cell("B2", fetched_data)
    set_cell("C2", comment)
    print("Step 3 complete -> cells written")


# --------------------------------------------------------------------------
# STEP 4 - SAVE + EXPORT, both as direct AppleScript commands
# --------------------------------------------------------------------------
def save_and_export(base_filename: str, xlsx_filename: str):
    """
    Saves the native .numbers file, then exports a .xlsx copy, both via
    Numbers' own `save` / `export` AppleScript commands -- no save
    dialogs, no keyboard navigation, no click coordinates.
    """
    print("Step 4: Saving native file and exporting to .xlsx...")

    numbers_path = os.path.join(OUTPUT_DIR, f"{base_filename}.numbers")
    xlsx_path = os.path.join(OUTPUT_DIR, xlsx_filename)

    save_script = f'''
    tell application "Numbers"
        save front document in POSIX file "{numbers_path}"
    end tell
    '''
    run_applescript(save_script)
    print(f"[ok] native file saved -> {numbers_path}")

    export_script = f'''
    tell application "Numbers"
        export front document to POSIX file "{xlsx_path}" as Microsoft Excel
    end tell
    '''
    run_applescript(export_script)
    print(f"[ok] xlsx exported -> {xlsx_path}")


# --------------------------------------------------------------------------
# STEP 5 - SCREENSHOT THE FINAL SHEET
# --------------------------------------------------------------------------
def take_screenshot(screenshot_filename: str):
    """
    Uses macOS's native `screencapture` CLI tool instead of
    pyautogui.screenshot(). pyautogui's screenshot function depends on
    pyscreeze -> Pillow, and on some Python versions (e.g. 3.14, which is
    very new) that dependency chain isn't installable yet, raising
    PyAutoGUIException. screencapture ships with macOS itself, so this
    sidesteps the dependency entirely and is arguably more reliable for
    Mac-only automation anyway.
    """
    print("Step 5: Taking screenshot of the final sheet...")
    activate_app("Numbers", wait_after=1.0)

    screenshot_path = os.path.join(OUTPUT_DIR, screenshot_filename)
    # -x  : no camera shutter sound
    result = subprocess.run(
        ["screencapture", "-x", screenshot_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"screencapture failed: {result.stderr.strip()}")

    print(f"Step 5 complete -> screenshot saved to {screenshot_path}")


# --------------------------------------------------------------------------
# MAIN ORCHESTRATION
# --------------------------------------------------------------------------
def main():
    print("=== Daily Status Report Bot starting ===")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    fetched_data = fetch_weather_via_browser()

    timestamp_str, fetched_data, comment, xlsx_filename, screenshot_filename = (
        build_report_row(fetched_data)
    )

    open_numbers_new_sheet()
    write_report_into_sheet(timestamp_str, fetched_data, comment)

    base_filename = xlsx_filename.replace(".xlsx", "")
    save_and_export(base_filename, xlsx_filename)

    take_screenshot(screenshot_filename)

    print("=== Daily Status Report Bot finished successfully ===")
    print(f"Row written: [{timestamp_str}] [{fetched_data}] [{comment}]")
    print(f"Files saved in: {OUTPUT_DIR}")


if __name__ == "__main__":
    print("Starting in 5 seconds -- move mouse to a screen corner to abort.")
    time.sleep(5)
    main()
