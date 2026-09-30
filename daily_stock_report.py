"""
Daily stock report bot  (pyautogui)

1. Opens the browser and goes to a public stock-price page
2. Copies the page text and extracts the stock price
3. Opens Microsoft Excel and creates a new workbook
4. Adds a row: date & time, fetched price, short comment
5. Saves the file as daily_report_YYYY-MM-DD.xlsx
6. Takes a screenshot of the final sheet

"""

import os
import re
import time
import subprocess
from datetime import datetime

import pyautogui
import pyperclip

pyautogui.FAILSAFE = True   # Move mouse to a screen corner to abort
pyautogui.PAUSE = 0.5       # Pause between actions

# ---------------- Settings (change as needed) ----------------
BROWSER = "safari"                     # or "google chrome"
STOCK_URL = "https://www.google.com/finance/quote/AAPL:NASDAQ"
STOCK_NAME = "Apple (AAPL)"
PAGE_KEYWORD = "Apple Inc"             # text that appears just before the price on the page
COMMENT = "Good for longterm investment"
SCREENSHOT_FOLDER = os.path.expanduser("~/Desktop")
# -------------------------------------------------------------

now = datetime.now()
timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
date_str = now.strftime("%Y-%m-%d")
file_name = f"daily_report_{date_str}"                       # Excel adds .xlsx
screenshot_path = os.path.join(SCREENSHOT_FOLDER, f"{file_name}.png")


def open_app(app_name, wait_seconds=5):
    """Open an app through Spotlight search."""
    pyautogui.hotkey('command', 'space', interval=0.1)
    time.sleep(1.5)
    pyautogui.write(app_name, interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(wait_seconds)


def extract_price(page_text, keyword):
    """Find the first $ price that appears after the company name."""
    start = page_text.find(keyword)
    search_area = page_text[start:] if start != -1 else page_text
    match = re.search(r'\$\s?([\d,]+\.\d{2})', search_area)
    return match.group(1).replace(',', '') if match else None


# ---------------- Step 1: Open the browser ----------------
print("Step 1: Open the browser")
open_app(BROWSER, wait_seconds=5)

# ---------------- Step 2: Open the stock page ----------------
print("Step 2: Open the stock website")
pyautogui.hotkey('command', 't', interval=0.1)    # New tab (address bar gets focus)
time.sleep(1)
pyautogui.hotkey('command', 'l', interval=0.1)    # Make sure address bar is focused
pyautogui.write(STOCK_URL, interval=0.05)
pyautogui.press('enter')
time.sleep(7)                                     # Wait for the page to load

# ---------------- Step 3: Copy the page and extract the price ----------------
print("Step 3: Copy the stock value")
pyperclip.copy("")                                # Clear clipboard first
pyautogui.hotkey('command', 'a', interval=0.1)    # Select all page text
time.sleep(1)
pyautogui.hotkey('command', 'c', interval=0.1)    # Copy
time.sleep(1.5)

page_text = pyperclip.paste()
price = extract_price(page_text, PAGE_KEYWORD)

if price:
    print(f"   Price found: {price}")
else:
    price = "Not found"
    debug_file = os.path.join(SCREENSHOT_FOLDER, "copied_page_text.txt")
    with open(debug_file, "w") as f:
        f.write(page_text)
    print(f"   Price not found. Copied text saved to {debug_file} for checking.")

# ---------------- Step 4: Open Excel and add the row ----------------
print("Step 4: Open Excel and add the data")
open_app("Microsoft Excel", wait_seconds=8)
pyautogui.hotkey('command', 'n', interval=0.1)    # New blank workbook
time.sleep(4)

# Header row (starts in cell A1)
pyautogui.write("Date & Time", interval=0.03)
pyautogui.press('tab')
pyautogui.write("Stock", interval=0.03)
pyautogui.press('tab')
pyautogui.write("Price (USD)", interval=0.03)
pyautogui.press('tab')
pyautogui.write("Comment", interval=0.03)
pyautogui.press('enter')                          # Moves back to column A, next row

# Data row
pyautogui.write(timestamp, interval=0.03)
pyautogui.press('tab')
pyautogui.write(STOCK_NAME, interval=0.03)
pyautogui.press('tab')
pyautogui.write(price, interval=0.03)
pyautogui.press('tab')
pyautogui.write(COMMENT, interval=0.03)
pyautogui.press('enter')
time.sleep(1)

# Widen columns so the screenshot is readable: select all, then AutoFit column width
pyautogui.hotkey('command', 'a', interval=0.1)
pyautogui.hotkey('ctrl', 'option', '0', interval=0.1)  # AutoFit may vary by Excel version; harmless if ignored
pyautogui.press('up')                                  # Deselect
time.sleep(1)

# ---------------- Step 5: Save with today's date ----------------
print("Step 5: Save the Excel file")
pyautogui.hotkey('command', 's', interval=0.1)
time.sleep(3)
pyautogui.hotkey('command', 'a', interval=0.1)    # Clear the default name in the Save dialog
pyautogui.write(file_name, interval=0.05)
time.sleep(1)
pyautogui.press('enter')
time.sleep(3)
# If the file already exists (running twice in a day), Excel asks to Replace
pyautogui.press('enter')
time.sleep(2)
print(f"   Saved as {file_name}.xlsx")

# ---------------- Step 6: Screenshot ----------------
print("Step 6: Take a screenshot")
time.sleep(1)
subprocess.run(["screencapture", "-x", screenshot_path])   # -x = no shutter sound
print(f"   Screenshot saved: {screenshot_path}")

print("Done!")