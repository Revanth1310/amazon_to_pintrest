import ctypes
import os
import re
import subprocess
import sys
import time
from pathlib import Path

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# =========================
# CONFIGURATION
# =========================

AMAZON_FILE = "data/amazon_products.xlsx"
POSTED_FILE = "data/posted_pins.xlsx"
IMAGE_DIR = "images"
CHROME_PROFILE_DIR = r"D:\AutomationProfile"
PINTEREST_URL = "https://in.pinterest.com/pin-creation-tool/"
AUTH_FILE = "auth_p.json"
HEADLESS = True

os.makedirs(IMAGE_DIR, exist_ok=True)

POSTED_COLUMNS = [
    "Category",
    "Name",
    "Picture",
    "Affiliate Link",
    "Description",
    "Status",
]


# =========================
# SYSTEM & UTILS
# =========================


def remove_non_bmp(text):
    if text is None:
        return ""
    return "".join(ch for ch in str(text) if ord(ch) <= 0xFFFF)


# =========================
# DATA MANAGEMENT
# =========================


def create_posted_file():
    if not os.path.exists(POSTED_FILE):
        os.makedirs(os.path.dirname(POSTED_FILE), exist_ok=True)
        pd.DataFrame(columns=POSTED_COLUMNS).to_excel(POSTED_FILE, index=False)
        print("Created posted_pins.xlsx")


def get_posted_product_names():
    create_posted_file()
    df = pd.read_excel(POSTED_FILE)
    if "Name" not in df.columns:
        return set()
    return set(df["Name"].dropna().astype(str).str.strip().str.lower())


def mark_as_posted(product):
    df = pd.read_excel(POSTED_FILE)
    new_row = {
        "Category": product["Category"],
        "Name": product["Name"],
        "Picture": product["Picture"],
        "Affiliate Link": product["Affiliate Link"],
        "Description": product["Description"],
        "Status": "Posted",
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(POSTED_FILE, index=False)
    print(f"Added to posted_pins.xlsx: {product['Name']}")


def is_already_posted(product_name):
    """Check posted_pins.xlsx immediately before starting a pin."""
    try:
        create_posted_file()
        df = pd.read_excel(POSTED_FILE)

        if "Name" not in df.columns:
            return False

        target = str(product_name).strip().lower()
        posted_names = set(
            df["Name"].dropna().astype(str).str.strip().str.lower()
        )
        return target in posted_names

    except Exception as e:
        print(f"Could not check posted_pins.xlsx: {e}")
        return False


def remove_from_amazon_products(product_name):
    try:
        df = pd.read_excel(AMAZON_FILE)
        target = product_name.strip().lower()
        df = df[df["Name"].astype(str).str.strip().str.lower() != target]
        df.to_excel(AMAZON_FILE, index=False)
        print(f"Removed from Amazon list: {product_name}")
    except Exception as e:
        print(f"Failed to remove product from Excel: {e}")



from PIL import Image


def download_image(url, product_name):
    if pd.isna(url):
        return None

    url = str(url).strip()

    if not url:
        return None

    try:
        safe_name = "".join(
            c for c in product_name
            if c.isalnum() or c in (" ", "_", "-")
        )[:100]

        # Save all processed images as JPG
        image_path = os.path.join(
            IMAGE_DIR,
            safe_name + ".jpg"
        )

        # If resized image already exists, reuse it
        if os.path.exists(image_path):
            return image_path

        # Download image
        response = requests.get(
            url,
            timeout=30,
            stream=True
        )
        response.raise_for_status()

        # Temporary original image
        temp_path = os.path.join(
            IMAGE_DIR,
            safe_name + "_original"
        )

        with open(temp_path, "wb") as f:
            for chunk in response.iter_content(1024):
                if chunk:
                    f.write(chunk)

        # Open downloaded image
        image = Image.open(temp_path)

        # Convert to RGB so it can always be saved as JPG
        if image.mode != "RGB":
            image = image.convert("RGB")

        # Resize to exactly 200 x 300 pixels
        image = image.resize(
            (200, 300),
            Image.Resampling.LANCZOS
        )

        # Save resized image
        image.save(
            image_path,
            "JPEG",
            quality=95
        )

        # Remove temporary original
        os.remove(temp_path)

        print(f"Downloaded and resized image: {product_name}")
        print(f"Size: 200 x 300 pixels")

        return image_path

    except Exception as e:
        print(
            f"Image download/resize failed for "
            f"{product_name}: {e}"
        )

        # Clean up temporary file if it exists
        try:
            if "temp_path" in locals() and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass

        return None




def prepare_pins():
    create_posted_file()
    posted_names = get_posted_product_names()
    amazon_df = pd.read_excel(AMAZON_FILE)

    pins_to_post = []

    for _, row in amazon_df.iterrows():
        product_name = str(row["Name"]).strip()

        # Skip products that are already posted
        # or don't have required information
        if (
            product_name.lower() in posted_names
            or pd.isna(row["Picture"])
            or pd.isna(row["Affiliate Link"])
        ):
            continue

        image_path = download_image(row["Picture"], product_name)

        if not image_path:
            continue

        pin_data = {
            "Category": row.get("Category", "General"),
            "Name": product_name,
            "Picture": row["Picture"],
            "Affiliate Link": row["Affiliate Link"],
            "Description": (
                ""
                if pd.isna(row.get("Description"))
                else str(row["Description"])
            ),
            "Image Path": image_path,
        }

        pins_to_post.append(pin_data)

    return pins_to_post


# =========================
# BROWSER AUTOMATION
# =========================


def start_browser():
    """Start Playwright Chromium using auth.json."""
    if not os.path.exists(AUTH_FILE):
        raise FileNotFoundError(f"Authentication file not found: {AUTH_FILE}")

    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=HEADLESS)

    context = browser.new_context(
        storage_state=AUTH_FILE,
        viewport={"width": 1920, "height": 1080},
    )

    page = context.new_page()

    print("Playwright browser started.")
    print(f"Authentication state loaded: {AUTH_FILE}")

    return playwright, browser, context, page


def post_pin(page, image_path, title, description, link):
    try:
        page.goto(PINTEREST_URL, wait_until="domcontentloaded", timeout=60000)

        file_input = page.locator("#storyboard-upload-input")
        file_input.wait_for(state="attached", timeout=60000)
        file_input.set_input_files(os.path.abspath(image_path))
        print("Image selected")

        try:
            file_input.wait_for(state="detached", timeout=60000)
        except PlaywrightTimeoutError:
            print("Warning: upload input did not disappear, continuing...")

        print("Image uploaded successfully")

        title = remove_non_bmp(title)
        description = remove_non_bmp(description)

        title_box = page.locator("#storyboard-selector-title")
        title_box.wait_for(state="visible", timeout=60000)
        title_box.fill(title)
        print("Title filled")

        desc_editor = page.locator(".public-DraftEditor-content")
        desc_editor.wait_for(state="visible", timeout=60000)
        desc_editor.click()
        desc_editor.fill(description[:800])
        print("Description filled")

        link_box = page.locator("#WebsiteField")
        link_box.wait_for(state="visible", timeout=60000)
        link_box.fill(str(link))

        page.keyboard.press("Tab")
        page.wait_for_timeout(1500)

        print("Link filled")

        return True

    except Exception as e:
        print(f"Draft creation failed: {e}")
        return False


def publish_pin(page):
    """
    Publish the current Pinterest pin.

    Pinterest initially renders the Done/Publish button as disabled.
    We wait specifically for it to become enabled instead of calling
    click() on a disabled locator.
    """
    try:
        print("Waiting for Pinterest to finish saving...")

        saved_status = page.locator(
            '[data-test-id="saving-status-saved"]'
        )
        saved_status.wait_for(state="visible", timeout=120000)

        print("Changes saved internally by Pinterest.")

        publish_container = page.locator(
            '[data-test-id="storyboard-creation-nav-done"]'
        )
        publish_container.wait_for(state="visible", timeout=60000)

        publish_btn = publish_container.locator("button")
        publish_btn.wait_for(state="visible", timeout=60000)

        # Give Pinterest's client-side validation a moment to update.
        page.wait_for_timeout(2000)

        # Wait for the actual button to become enabled.
        # Do not use force=True because Pinterest can reject an invalid
        # or incomplete pin when the button is genuinely disabled.
        print("Waiting for Publish button to become enabled...")

        deadline = time.time() + 90

        while time.time() < deadline:
            try:
                if publish_btn.is_enabled():
                    print("Publish button is enabled.")
                    break

                disabled = publish_btn.get_attribute("disabled")
                aria_disabled = publish_btn.get_attribute("aria-disabled")

                print(
                    f"Publish button still disabled "
                    f"(disabled={disabled}, aria-disabled={aria_disabled})"
                )

                # Re-trigger validation by moving focus away from the
                # currently active form element.
                try:
                    page.keyboard.press("Tab")
                except Exception:
                    pass

                page.wait_for_timeout(2000)

            except Exception as e:
                print(f"Publish button check failed: {e}")
                page.wait_for_timeout(1000)

        if not publish_btn.is_enabled():
            print("Publish button did not become enabled.")
            print("Pinterest is still rejecting the pin form.")

            # Print useful diagnostic information without dumping the
            # entire page source.
            try:
                print(f"Current URL: {page.url}")
                print(
                    "Publish button disabled attribute:",
                    publish_btn.get_attribute("disabled"),
                )
                print(
                    "Publish button aria-disabled:",
                    publish_btn.get_attribute("aria-disabled"),
                )
                print(
                    "Publish button text:",
                    publish_btn.inner_text(timeout=5000),
                )
            except Exception as diag_error:
                print(f"Could not collect publish diagnostics: {diag_error}")

            return False

        publish_btn.scroll_into_view_if_needed()
        page.wait_for_timeout(1000)

        # Click once. The previous Selenium code clicked twice, but
        # double-clicking can cause duplicate/unstable Pinterest actions.
        # Click Publish/Done twice, as requested.
        publish_btn.click(timeout=30000)
        print("Publish button clicked (1/2)")

        page.wait_for_timeout(1000)

        try:
            if publish_btn.is_visible() and publish_btn.is_enabled():
                publish_btn.click(timeout=30000)
                print("Publish button clicked (2/2)")
            else:
                print("Second publish click skipped: button is no longer available.")
        except Exception as second_click_error:
            print(f"Second publish click skipped: {second_click_error}")

        # Give Pinterest time to navigate/update the creation flow.
        page.wait_for_timeout(3000)

        return True

    except Exception as e:
        print(f"Publish failed: {e}")
        return False



# =========================
# MAIN EXECUTION
# =========================

if __name__ == "__main__":
    playwright = None
    browser = None
    context = None
    page = None

    try:
        print("\nReading all available pins...")

        pins = prepare_pins()

        print(f"\nPins Ready: {len(pins)}\n")

        if not pins:
            print("No pins available to post.")
            print("Stopping program.")
            sys.exit(0)

        playwright, browser, context, page = start_browser()

        for i, pin in enumerate(pins, start=1):
            print(f"\nPosting {i}/{len(pins)}: {pin['Name']}")


            # Re-check posted_pins.xlsx immediately before posting.
            if is_already_posted(pin["Name"]):
                print(f"Already posted according to {POSTED_FILE}: {pin['Name']}")
                print("Skipping this pin.")
                print("=========================================")
                continue

            clean_title = pin["Name"].split("|")[0].strip()

            if len(clean_title) > 80 and "," in clean_title:
                clean_title = clean_title.split(",")[0].strip()

            if len(clean_title) > 80:
                clean_title = clean_title[:80].rsplit(" ", 1)[0]

            clean_desc = str(pin.get("Description", ""))

            if clean_desc.lower() == "nan":
                clean_desc = ""

            draft_success = post_pin(
                page,
                pin["Image Path"],
                clean_title,
                clean_desc,
                pin["Affiliate Link"],
            )

            if draft_success:
                print("Draft created successfully. Publishing...")
                is_published = publish_pin(page)

                if is_published:
                    mark_as_posted(pin)
                    remove_from_amazon_products(pin["Name"])
                    print(f"Successfully posted: {pin['Name']}")
                else:
                    print(f"Failed to publish pin for: {pin['Name']}")

            if i < len(pins):
                time.sleep(60)

            print("=========================================")

    except Exception as e:
        print(f"An error occurred during execution: {e}")

    finally:
        print("\nClosing Playwright browser...")

        try:
            if page:
                page.close()
        except BaseException:
            pass

        try:
            if context:
                context.close()
        except BaseException:
            pass

        try:
            if browser:
                browser.close()
        except BaseException:
            pass

        try:
            if playwright:
                playwright.stop()
        except BaseException:
            pass

        print("\nDone!")
