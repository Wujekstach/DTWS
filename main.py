import os
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
import sys
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import win32com.client as win32
from dotenv import set_key

load_dotenv()

# Configuration
WEBSITE_URL = os.getenv("WEBSITE_URL")
EMAIL_RECIPIENTS = [
    email.strip()
    for email in os.getenv("EMAIL_RECIPIENTS", "").split(",")
    if email.strip()
]


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
    
SCREENSHOT_DIR = BASE_DIR / "screenshots"
SCREENSHOT_DIR.mkdir(exist_ok=True)

BROWSER_PROFILE = BASE_DIR / "browser_profile"


def check_subscription_emails():
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # Inbox
    
    messages = inbox.Items
    cutoff = datetime.now() - timedelta(hours=24)
    filter_date = cutoff.strftime("%m/%d/%Y %I:%M %p")
    messages = messages.Restrict(
        f"[ReceivedTime] >= '{filter_date}'"
    )

    messages.Sort("[ReceivedTime]", True)

    recipients = [
        email.strip().lower()
        for email in os.getenv("EMAIL_RECIPIENTS", "").split(",")
        if email.strip()
    ]
    
    ALLOWED_DOMAINS = [
        "@t-mobile.pl",
        "@telekom.de"
    ]
    changed = False

    for message in messages:
        try:
            if message.Class != 43:  # MailItem
                continue

            if message.UnRead is False:
                continue

            subject = message.Subject.strip().upper()
            sender = message.SenderEmailAddress.strip().lower()

            if not any(sender.endswith(domain) for domain in ALLOWED_DOMAINS):
                print(f"Ignoring unauthorized sender: {sender}")
                continue

            if subject == "SUBSCRIBE":
                if sender not in recipients:
                    recipients.append(sender)
                    changed = True
                    print(f"Subscribed: {sender}")

            elif subject == "UNSUBSCRIBE":
                if sender in recipients:
                    recipients.remove(sender)
                    changed = True
                    print(f"Unsubscribed: {sender}")

            else:
                continue

            message.UnRead = False

        except Exception as e:
            print(f"Error processing email: {e}")

    if changed:
        set_key(
            str(BASE_DIR / ".env"),
            "EMAIL_RECIPIENTS",
            ",".join(recipients)
        )

        os.environ["EMAIL_RECIPIENTS"] = ",".join(recipients)

        print("Mailing list updated.")

    return recipients

def take_screenshot():
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    screenshot_path = SCREENSHOT_DIR / f"screenshot_{timestamp}.png"
    user_data_dir=str(BROWSER_PROFILE)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(BASE_DIR / "browser_profile"),
            headless=False,
            channel="chrome",
            viewport={"width": 1920, "height": 1080}
        )

        page = context.pages[0] if context.pages else context.new_page()

        page.goto(WEBSITE_URL, wait_until="domcontentloaded")

        # Wait for manual login if necessary
        print("Waiting for website to load...")

        page.wait_for_load_state("networkidle", timeout=60000)
        # Close optional banner if present
        close_button = page.locator('button.close[title="Zamknij"]')

        if close_button.is_visible():
            close_button.click()
            print("Closed notification banner")
        else:
            print("No notification banner found")
            
        page.screenshot(
            path=str(screenshot_path),
            full_page=True
        )

        context.close()

    return screenshot_path


def send_email(screenshot_path):
    if not EMAIL_RECIPIENTS:
        raise ValueError("No email recipients configured in .env")

    outlook = win32.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)

    mail.To = "; ".join(EMAIL_RECIPIENTS)
    mail.Subject = f"Daily T-Wall Screenshot - {datetime.now():%Y-%m-%d}"

    mail.Body = (
        f"Daily screenshot of {WEBSITE_URL}\n"
        f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}"
    )

    mail.Attachments.Add(str(screenshot_path.resolve()))

    mail.Send()

    print("Email sent successfully!")


def cleanup_old_screenshots(keep=10):
    screenshots = sorted(
        SCREENSHOT_DIR.glob("screenshot_*.png"),
        key=lambda file: file.stat().st_mtime,
        reverse=True
    )

    for old_file in screenshots[keep:]:
        old_file.unlink()
        print(f"Deleted old screenshot: {old_file.name}")

def main():
    # Skip weekends (Saturday = 5, Sunday = 6)
    if datetime.today().weekday() >= 5:
        print("Weekend detected. Skipping execution.")
        return
    check_subscription_emails()

    screenshot = take_screenshot()
    send_email(screenshot)

    cleanup_old_screenshots(keep=5)


if __name__ == "__main__":
    main()
