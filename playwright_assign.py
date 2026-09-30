"""
WhatsApp Message Sender + Smart Data Extractor
Assignment: Playwright (Python) + WhatsApp Web + Excel
File: playwright_assign.py

Setup (run once in the VS Code terminal):
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python playwright_assign.py

First run: scan the QR code with your phone (WhatsApp > Linked devices).
The login is saved in the "wa_session" folder, so later runs skip the QR.
"""

import json
import os
import random
import sys
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ---------------- Settings ----------------
CONTACTS_FILE = "contacts.xlsx"
SESSION_DIR = "wa_session"          # keeps you logged in between runs
SCREENSHOT_DIR = "screenshots"
REPORT_JSON = "report.json"
REPORT_XLSX = "report.xlsx"
DEFAULT_TEMPLATE = "Hi {name}, hope you are doing well! This is an automated test message."

# WhatsApp Web selectors (WhatsApp changes these sometimes; update here if needed)
SEARCH_BOX = 'div[contenteditable="true"][data-tab="3"]'
MESSAGE_BOX = 'footer div[contenteditable="true"]'
CHAT_LIST_LOADED = '#pane-side'
SENT_TICK = ('div.message-out span[data-icon="msg-check"], '
             'div.message-out span[data-icon="msg-dblcheck"]')
INCOMING_TEXT = 'div.message-in span.selectable-text'


# ---------------- Helpers ----------------
def human_pause(page, low=2, high=5):
    """Random 2-5 second pause so the bot does not act like a robot."""
    page.wait_for_timeout(random.randint(low * 1000, high * 1000))


def create_sample_contacts():
    """Create a sample contacts.xlsx if the file does not exist yet."""
    sample = pd.DataFrame([
        {"Name": "Ravi", "Phone": "+919999999999", "Message": "Hello {name}, your order is ready!"},
        {"Name": "Priya", "Phone": "+918888888888", "Message": ""},
    ])
    sample.to_excel(CONTACTS_FILE, index=False)
    print(f"Created sample {CONTACTS_FILE}. Edit it with real contacts and run again.")


def read_contacts():
    df = pd.read_excel(CONTACTS_FILE, dtype=str).fillna("")
    contacts = []
    for _, row in df.iterrows():
        contacts.append({
            "name": row.get("Name", "").strip(),
            "phone": row.get("Phone", "").strip(),
            "template": row.get("Message", "").strip() or DEFAULT_TEMPLATE,
        })
    return contacts


def login(page):
    """Open WhatsApp Web and wait until the chat list appears (QR scanned)."""
    page.goto("https://web.whatsapp.com")
    print("Waiting for WhatsApp Web to load. Scan the QR code if asked...")
    page.wait_for_selector(CHAT_LIST_LOADED, timeout=120000)  # 2 minutes to scan QR
    print("Logged in successfully.")
    human_pause(page)


def open_chat(page, search_text):
    """Search by name or number and open the first result. Returns True if chat opened."""
    search = page.wait_for_selector(SEARCH_BOX, timeout=15000)
    search.click()
    page.keyboard.press("Meta+A" if sys.platform == "darwin" else "Control+A")  # Cmd+A on Mac
    page.keyboard.press("Backspace")
    page.keyboard.type(search_text, delay=100)   # human-like typing
    page.wait_for_timeout(3000)                  # let search results load
    page.keyboard.press("Enter")
    try:
        page.wait_for_selector(MESSAGE_BOX, timeout=8000)
        return True
    except PlaywrightTimeout:
        return False


def open_chat_by_url(page, phone):
    """Fallback: open chat directly with the phone number (works for unsaved numbers)."""
    number = phone.replace("+", "").replace(" ", "")
    page.goto(f"https://web.whatsapp.com/send?phone={number}")
    try:
        page.wait_for_selector(MESSAGE_BOX, timeout=30000)
        return True
    except PlaywrightTimeout:
        return False


def send_message(page, text):
    """Type the message in the 'Type a message' box and confirm it was sent."""
    ticks_before = len(page.query_selector_all(SENT_TICK))
    box = page.wait_for_selector(MESSAGE_BOX, timeout=10000)
    box.click()
    page.keyboard.type(text.replace("\n", " "), delay=50)
    human_pause(page, 1, 2)
    page.keyboard.press("Enter")

    # Confirm: wait until a new sent tick (single or double) appears
    for _ in range(20):  # up to ~20 seconds
        page.wait_for_timeout(1000)
        if len(page.query_selector_all(SENT_TICK)) > ticks_before:
            return True
    return False


def extract_last_messages(page, count=3):
    """Extract the last N messages received from the contact in the open chat."""
    try:
        page.wait_for_selector("div.message-in, div.message-out", timeout=10000)
    except PlaywrightTimeout:
        return []
    elements = page.query_selector_all(INCOMING_TEXT)
    return [el.inner_text().strip() for el in elements[-count:]]


def save_report(results):
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=4, ensure_ascii=False)

    excel_rows = []
    for r in results:
        row = dict(r)
        row["last_messages"] = " | ".join(r["last_messages"])
        excel_rows.append(row)
    pd.DataFrame(excel_rows).to_excel(REPORT_XLSX, index=False)
    print(f"Report saved: {REPORT_JSON} and {REPORT_XLSX}")


# ---------------- Main flow ----------------
def main():
    if not os.path.exists(CONTACTS_FILE):
        create_sample_contacts()
        return

    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    contacts = read_contacts()
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            SESSION_DIR, headless=False, viewport={"width": 1280, "height": 850}
        )
        page = browser.pages[0] if browser.pages else browser.new_page()

        try:
            login(page)
        except PlaywrightTimeout:
            print("Login timed out. QR code was not scanned in time.")
            browser.close()
            return

        for c in contacts:
            print(f"\nProcessing {c['name']} ({c['phone']})")
            result = {
                "name": c["name"],
                "phone": c["phone"],
                "message": "",
                "status": "Not sent",
                "screenshot": "",
                "last_messages": [],
                "error": "",
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            try:
                # 1. Search contact: by number first, then by name, then direct URL
                opened = open_chat(page, c["phone"].replace("+", "")) if c["phone"] else False
                if not opened and c["name"]:
                    opened = open_chat(page, c["name"])
                if not opened and c["phone"]:
                    opened = open_chat_by_url(page, c["phone"])
                if not opened:
                    raise Exception("Contact not found")
                human_pause(page)

                # 2. Personalize and send
                text = c["template"].replace("{name}", c["name"])
                result["message"] = text
                if send_message(page, text):
                    result["status"] = "Sent"
                    print("  Message sent.")
                else:
                    result["status"] = "Not confirmed"
                    print("  Message typed, but sent tick not seen.")

                # 3. Screenshot of the sent message
                shot = os.path.join(SCREENSHOT_DIR, f"{c['name'] or 'contact'}_{datetime.now():%H%M%S}.png")
                page.screenshot(path=shot)
                result["screenshot"] = shot

                # 4. Smart extraction: last 3 messages from this contact
                human_pause(page)
                result["last_messages"] = extract_last_messages(page, 3)
                print(f"  Last messages: {result['last_messages']}")

            except Exception as e:
                result["status"] = "Failed"
                result["error"] = str(e)
                print(f"  Error: {e}")
                page.keyboard.press("Escape")  # close any open search/popup

            results.append(result)
            human_pause(page)  # random 2-5 sec gap before the next contact

        browser.close()

    save_report(results)
    print("\nDone.")


if __name__ == "__main__":
    main()