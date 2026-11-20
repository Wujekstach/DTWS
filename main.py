
import os
import sys
import json
import logging
import hashlib
import math

from datetime import datetime, timedelta
from pathlib import Path

import win32com.client as win32

from dotenv import load_dotenv
from PIL import Image, ImageDraw

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError
)


# ============================================================
# CONFIGURATION
# ============================================================

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

ENV_FILE = BASE_DIR / ".env"
load_dotenv(ENV_FILE)

WEBSITE_URL = os.getenv("WEBSITE_URL", "").strip()

SCREENSHOT_DIR = BASE_DIR / "screenshots"
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

BROWSER_PROFILE = BASE_DIR / "browser_profile"

SUBSCRIPTIONS_FILE = BASE_DIR / "subscriptions.json"

SCREENSHOT_HEIGHT = 2750
VIEWPORT_WIDTH = 1920
VIEWPORT_HEIGHT = 1080

ALLOWED_DOMAINS = (
    "@t-mobile.pl",
    "@telekom.de",
)

OUTLOOK_INBOX = 6

# Arrow appearance
ARROW_COLOR = (230, 0, 0)
ARROW_OUTLINE = (255, 255, 255)
ARROW_LINE_WIDTH = 12
ARROW_OUTLINE_WIDTH = 22
ARROW_HEAD_LENGTH = 35
ARROW_HEAD_WIDTH = 20


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(
            BASE_DIR / "automation.log",
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


# ============================================================
# SUBSCRIPTION STORAGE
# ============================================================

def load_subscriptions():
    """
    Load subscriber email addresses and their keywords.

    Format:
    {
        "user@example.com": [
            "System maintenance",
            "Network outage"
        ]
    }
    """

    if not SUBSCRIPTIONS_FILE.exists():
        return {}

    try:
        with open(SUBSCRIPTIONS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)

        return {
            email.lower(): keywords
            for email, keywords in data.items()
            if isinstance(keywords, list)
        }

    except Exception:
        logger.exception("Could not load subscriptions.json")
        return {}


def save_subscriptions(subscriptions):
    """Save subscriptions to JSON."""

    temp_file = SUBSCRIPTIONS_FILE.with_suffix(".tmp")

    with open(temp_file, "w", encoding="utf-8") as file:
        json.dump(
            subscriptions,
            file,
            indent=4,
            ensure_ascii=False
        )

    temp_file.replace(SUBSCRIPTIONS_FILE)


def parse_keywords(body):
    """
    Extract keywords separated by semicolons or new lines.
    Duplicate keywords are removed.
    """

    keywords = []

    for line in body.splitlines():
        for phrase in line.split(";"):
            phrase = phrase.strip()

            if phrase and phrase.lower() not in [
                keyword.lower() for keyword in keywords
            ]:
                keywords.append(phrase)

    return keywords


# ============================================================
# OUTLOOK SENDER ADDRESS
# ============================================================

def get_sender_email(message):
    """Get SMTP address, including Exchange accounts."""

    sender_email = ""

    try:
        sender_email = message.SenderEmailAddress or ""

        if sender_email.upper().startswith("/O="):
            try:
                exchange_user = message.Sender.GetExchangeUser()

                if exchange_user:
                    sender_email = (
                        exchange_user.PrimarySmtpAddress or sender_email
                    )

            except Exception:
                pass

        if not sender_email or sender_email.upper().startswith("/O="):
            try:
                sender_email = message.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
                )
            except Exception:
                pass

    except Exception:
        logger.exception("Could not determine sender address")

    return sender_email.strip().lower()


# ============================================================
# SUBSCRIPTION EMAIL PROCESSING
# ============================================================

def check_subscription_emails():
    """
    Process unread Outlook messages received within the last 24 hours.

    SUBSCRIBE:
        Adds a subscriber and their keywords.

    KEYWORDS:
        Replaces the subscriber's keyword list.

    UNSUBSCRIBE:
        Removes the subscriber.
    """

    logger.info("Checking subscription emails...")

    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = namespace.GetDefaultFolder(OUTLOOK_INBOX)

    messages = inbox.Items.Restrict("[Unread] = True")
    messages.Sort("[ReceivedTime]", True)

    # Build a snapshot before modifying message read status
    pending_messages = []

    message = messages.GetFirst()

    while message:
        pending_messages.append(message)
        message = messages.GetNext()

    subscriptions = load_subscriptions()

    cutoff_time = datetime.now() - timedelta(hours=24)

    for message in pending_messages:

        try:
            received_time = message.ReceivedTime.replace(tzinfo=None)

            if received_time < cutoff_time:
                continue

            subject = (message.Subject or "").strip().upper()

            if subject not in (
                "SUBSCRIBE",
                "KEYWORDS",
                "UNSUBSCRIBE"
            ):
                continue

            sender_email = get_sender_email(message)

            if not sender_email:
                continue

            # Validate sender domain
            if not any(
                sender_email.endswith(domain)
                for domain in ALLOWED_DOMAINS
            ):
                logger.warning(
                    "Rejected subscription request from %s",
                    sender_email
                )
                continue

            if subject == "UNSUBSCRIBE":

                if sender_email in subscriptions:
                    del subscriptions[sender_email]
                    logger.info("Unsubscribed: %s", sender_email)
                else:
                    logger.info(
                        "Not subscribed: %s",
                        sender_email
                    )

            else:

                keywords = parse_keywords(message.Body or "")

                if not keywords:
                    logger.warning(
                        "%s request from %s contains no keywords",
                        subject,
                        sender_email
                    )
                    continue

                if subject == "SUBSCRIBE":

                    subscriptions[sender_email] = keywords

                    logger.info(
                        "Subscribed %s with %s keywords",
                        sender_email,
                        len(keywords)
                    )

                elif subject == "KEYWORDS":

                    if sender_email not in subscriptions:
                        logger.warning(
                            "KEYWORDS request from non-subscriber: %s",
                            sender_email
                        )
                        continue

                    subscriptions[sender_email] = keywords

                    logger.info(
                        "Updated keywords for %s: %s",
                        sender_email,
                        len(keywords)
                    )

            message.UnRead = False
            message.Save()

        except Exception:
            logger.exception("Error processing subscription message")

    save_subscriptions(subscriptions)

    logger.info(
        "Subscription processing finished. Total subscribers: %s",
        len(subscriptions)
    )

    return subscriptions


# ============================================================
# FIND KEYWORDS ON WEBSITE
# ============================================================

def find_keyword_positions(page, keywords):
    """
    Find all occurrences of each keyword in website text.

    Returns positions in CSS pixels relative to the full page.
    """

    javascript = """
    (keywords) => {

        const results = {};

        const walker = document.createTreeWalker(
            document.body,
            NodeFilter.SHOW_TEXT
        );

        let node;

        while (node = walker.nextNode()) {

            const parent = node.parentElement;

            if (!parent) continue;

            const tag = parent.tagName.toLowerCase();

            if ([
                "script",
                "style",
                "noscript",
                "svg"
            ].includes(tag)) continue;

            const style = window.getComputedStyle(parent);

            if (
                style.display === "none" ||
                style.visibility === "hidden"
            ) continue;

            const text = node.nodeValue.toLowerCase();

            for (const keyword of keywords) {

                const search = keyword.toLowerCase();

                if (!search) continue;

                let index = 0;

                while (
                    (index = text.indexOf(search, index)) !== -1
                ) {

                    const range = document.createRange();

                    range.setStart(node, index);
                    range.setEnd(node, index + search.length);

                    const rects = range.getClientRects();

                    if (!results[keyword]) {
                        results[keyword] = [];
                    }

                    if (rects.length > 0) {

                        const rect = rects[0];

                        results[keyword].push({
                            x: rect.left,
                            y: rect.top + window.scrollY,
                            width: rect.width,
                            height: rect.height
                        });
                    }

                    index += search.length;
                }
            }
        }

        return results;
    }
    """

    try:
        return page.evaluate(javascript, keywords)

    except Exception:
        logger.exception("Could not find keyword positions")
        return {}


# ============================================================
# TAKE WEBSITE SCREENSHOT
# ============================================================

def take_screenshot(all_keywords):
    """
    Open website using persistent Chrome profile.

    Capture full page, find all requested keywords,
    then crop screenshot to top 2750 pixels.

    Returns:
        screenshot_path
        keyword_positions
        scale_factor
    """

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    screenshot_path = (
        SCREENSHOT_DIR / f"screenshot_{timestamp}_base.png"
    )

    logger.info("Starting browser...")

    with sync_playwright() as playwright:

        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE),
            channel="chrome",
            headless=False,
            viewport={
                "width": VIEWPORT_WIDTH,
                "height": VIEWPORT_HEIGHT
            }
        )

        try:
            page = (
                context.pages[0]
                if context.pages
                else context.new_page()
            )

            logger.info("Opening %s", WEBSITE_URL)

            try:
                page.goto(
                    WEBSITE_URL,
                    wait_until="networkidle",
                    timeout=60000
                )

            except PlaywrightTimeoutError:
                logger.warning("Network idle timeout")

                page.wait_for_load_state(
                    "domcontentloaded",
                    timeout=15000
                )

            page.wait_for_timeout(3000)

            # Close banner if present
            close_button = page.locator(
                'button.close[title="Zamknij"]'
            )

            try:
                if close_button.is_visible(timeout=3000):
                    close_button.click()
                    logger.info("Closed website banner")

            except Exception:
                pass

            # Find keywords before taking screenshot
            keyword_positions = find_keyword_positions(
                page,
                all_keywords
            )

            viewport_width = page.evaluate("window.innerWidth")

            # Capture full page
            logger.info("Taking screenshot...")

            page.screenshot(
                path=str(screenshot_path),
                full_page=True
            )

            # Crop screenshot
            with Image.open(screenshot_path) as image:

                scale_factor = image.width / viewport_width

                crop_height = min(
                    SCREENSHOT_HEIGHT,
                    image.height
                )

                cropped_image = image.crop(
                    (0, 0, image.width, crop_height)
                )

                cropped_image.save(screenshot_path)

                logger.info(
                    "Screenshot dimensions: %s x %s",
                    image.width,
                    crop_height
                )

            return (
                screenshot_path,
                keyword_positions,
                scale_factor
            )

        finally:
            context.close()
            logger.info("Browser closed")


# ============================================================
# DRAW ARROWS
# ============================================================

def draw_arrow(draw, target_x, target_y, image_width, image_height):
    """
    Draw a large red arrow pointing towards a matching phrase.
    White outline improves visibility against different backgrounds.
    """

    target_x = int(target_x)
    target_y = int(target_y)

    # Keep arrow within screenshot
    target_y = max(20, min(target_y, image_height - 20))

    # Arrow starts near the matching phrase
    if target_x > 160:
        start_x = max(25, target_x - 150)
        tip_x = target_x - 8
    else:
        start_x = min(image_width - 25, target_x + 150)
        tip_x = target_x + 8

    start = (start_x, target_y)
    tip = (tip_x, target_y)

    dx = tip[0] - start[0]
    dy = tip[1] - start[1]

    length = math.sqrt(dx * dx + dy * dy)

    if length == 0:
        return

    ux = dx / length
    uy = dy / length

    px = -uy
    py = ux

    base_x = tip_x - ux * ARROW_HEAD_LENGTH
    base_y = target_y - uy * ARROW_HEAD_LENGTH

    left = (
        base_x + px * ARROW_HEAD_WIDTH,
        base_y + py * ARROW_HEAD_WIDTH
    )

    right = (
        base_x - px * ARROW_HEAD_WIDTH,
        base_y - py * ARROW_HEAD_WIDTH
    )

    polygon = [
        tip,
        left,
        right
    ]

    # White outline
    draw.line(
        [start, tip],
        fill=ARROW_OUTLINE,
        width=ARROW_OUTLINE_WIDTH
    )

    draw.polygon(
        polygon,
        fill=ARROW_COLOR,
        outline=ARROW_OUTLINE,
        width=5
    )

    # Red arrow
    draw.line(
        [start, (base_x, base_y)],
        fill=ARROW_COLOR,
        width=ARROW_LINE_WIDTH
    )


# ============================================================
# CREATE PERSONALIZED SCREENSHOT
# ============================================================

def create_personalized_screenshot(
    original_path,
    keyword_positions,
    subscriber_keywords,
    scale_factor,
    subscriber_email,
    timestamp
):
    """
    Create a separate screenshot for each subscriber.

    Only their own matching keywords get arrows.
    """

    matches = []

    for keyword in subscriber_keywords:

        positions = keyword_positions.get(keyword, [])

        for position in positions:

            x = (
                position["x"] * scale_factor
            )

            y = (
                position["y"] * scale_factor
            )

            if y < SCREENSHOT_HEIGHT:
                matches.append({
                    "keyword": keyword,
                    "x": x,
                    "y": y,
                    "width": position["width"] * scale_factor,
                    "height": position["height"] * scale_factor
                })

    # Unique filename for each subscriber
    email_hash = hashlib.md5(
        subscriber_email.encode("utf-8")
    ).hexdigest()[:10]

    output_path = (
        SCREENSHOT_DIR /
        f"screenshot_{timestamp}_{email_hash}.png"
    )

    with Image.open(original_path) as original:

        image = original.convert("RGB")
        draw = ImageDraw.Draw(image)

        for match in matches:

            target_x = match["x"]

            target_y = (
                match["y"] +
                match["height"] / 2
            )

            draw_arrow(
                draw,
                target_x,
                target_y,
                image.width,
                image.height
            )

        image.save(output_path)

    logger.info(
        "%s: %s matching occurrences",
        subscriber_email,
        len(matches)
    )

    return output_path, matches


# ============================================================
# SEND INDIVIDUAL EMAIL
# ============================================================

def send_email(
    subscriber_email,
    screenshot_path,
    matches
):
    """
    Send an individual email to one subscriber.

    Each subscriber receives only their own matching alerts.
    """

    outlook = win32.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)

    mail.To = subscriber_email

    mail.Subject = "Daily Website Screenshot"

    mail.Body = (
        f"Daily screenshot of {WEBSITE_URL}\n"
        f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}\n\n"
    )

    if matches:

        mail.Body += "Your keyword matches:\n\n"

        # Summarize matches by keyword
        counts = {}

        for match in matches:
            keyword = match["keyword"]
            counts[keyword] = counts.get(keyword, 0) + 1

        for keyword, count in counts.items():

            mail.Body += (
                f"- {keyword}: {count} occurrence(s)\n"
            )

        mail.Body += (
            "\nRed arrows in the screenshot indicate "
            "the matching locations.\n"
        )

    else:

        mail.Body += (
            "No matches for your keywords were found today.\n"
        )

    mail.Attachments.Add(
        str(screenshot_path.resolve())
    )

    mail.Send()

    logger.info("Email sent to %s", subscriber_email)


# ============================================================
# CLEANUP OLD SCREENSHOTS
# ============================================================

def cleanup_old_screenshots(keep=5):
    """
    Keep the newest five screenshot runs.

    All personalized screenshots from a run are kept together.
    """

    screenshots = list(
        SCREENSHOT_DIR.glob("screenshot_*.png")
    )

    runs = {}

    for screenshot in screenshots:

        # Extract timestamp from filename
        parts = screenshot.stem.split("_")

        if len(parts) < 4:
            continue

        timestamp = "_".join(parts[1:3])

        runs.setdefault(timestamp, []).append(screenshot)

    sorted_runs = sorted(
        runs.items(),
        key=lambda item: item[0],
        reverse=True
    )

    for timestamp, files in sorted_runs[keep:]:

        for screenshot in files:

            try:
                screenshot.unlink()

                logger.info(
                    "Deleted old screenshot: %s",
                    screenshot.name
                )

            except Exception:
                logger.exception(
                    "Could not delete %s",
                    screenshot.name
                )


# ============================================================
# MAIN
# ============================================================

def main():

    logger.info("========== AUTOMATION STARTED ==========")

    # Skip weekends
    if datetime.today().weekday() >= 5:
        logger.info("Weekend detected. Skipping.")
        return

    if not WEBSITE_URL:
        logger.error("WEBSITE_URL missing from .env")
        return

    try:

        # Process subscription requests
        subscriptions = check_subscription_emails()

        if not subscriptions:
            logger.info("No subscribers. Nothing to send.")
            return

        # Collect all unique keywords
        all_keywords = list({
            keyword
            for keywords in subscriptions.values()
            for keyword in keywords
        })

        if not all_keywords:
            logger.info("No keywords configured.")
            return

        logger.info(
            "Monitoring %s unique keywords for %s subscribers",
            len(all_keywords),
            len(subscriptions)
        )

        # Take one screenshot and locate all keywords
        (
            original_path,
            keyword_positions,
            scale_factor
        ) = take_screenshot(all_keywords)

        timestamp = datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )

        # Create and send individual emails
        for email, keywords in subscriptions.items():

            try:

                personalized_path, matches = (
                    create_personalized_screenshot(
                        original_path,
                        keyword_positions,
                        keywords,
                        scale_factor,
                        email,
                        timestamp
                    )
                )

                send_email(
                    email,
                    personalized_path,
                    matches
                )

            except Exception:
                logger.exception(
                    "Failed to process subscriber %s",
                    email
                )

        # Remove old screenshot runs
        cleanup_old_screenshots(keep=5)

    except Exception:
        logger.exception("Automation failed")

    finally:
        logger.info("========== AUTOMATION FINISHED ==========")


if __name__ == "__main__":
    main()
