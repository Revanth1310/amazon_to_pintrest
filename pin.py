#"C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="D:\AutomationProfile"
import os
import pandas as pd
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import ctypes
import sys
import subprocess
import re
import time


# ---------------- Administrator Check ---------------- #

def is_admin():
    """Return True if the script is running as Administrator."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def run_command(cmd):
    """Run a command and raise an exception if it fails."""
    subprocess.run(cmd, shell=True, check=True)


def get_current_timeouts():
    """Returns (AC_timeout, DC_timeout) in seconds."""

    output = subprocess.check_output(
        "powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE",
        shell=True,
        text=True,
    )

    ac_match = re.search(
        r"Current AC Power Setting Index:\s+0x([0-9A-Fa-f]+)",
        output
    )

    dc_match = re.search(
        r"Current DC Power Setting Index:\s+0x([0-9A-Fa-f]+)",
        output
    )

    if not ac_match or not dc_match:
        raise Exception("Could not read display timeout values.")

    ac = int(ac_match.group(1), 16)
    dc = int(dc_match.group(1), 16)

    return ac, dc


def print_timeouts(title):
    ac, dc = get_current_timeouts()

    print(f"\n{'=' * 50}")
    print(title)
    print(f"{'=' * 50}")
    print(f"AC Timeout : {ac} seconds ({ac/60:.2f} minutes)")
    print(f"DC Timeout : {dc} seconds ({dc/60:.2f} minutes)")


def set_display_timeout(ac_seconds, dc_seconds):
    """Set display timeout (seconds)."""

    run_command(
        f"powercfg /setacvalueindex SCHEME_CURRENT SUB_VIDEO VIDEOIDLE {ac_seconds}"
    )

    run_command(
        f"powercfg /setdcvalueindex SCHEME_CURRENT SUB_VIDEO VIDEOIDLE {dc_seconds}"
    )

    run_command("powercfg /setactive SCHEME_CURRENT")
PINTEREST_URL = "https://in.pinterest.com/pin-creation-tool/"

# =========================
# CONFIG
# =========================

AMAZON_FILE = "data/amazon_products.xlsx"
POSTED_FILE = "data/posted_pins.xlsx"
IMAGE_DIR = "images"

os.makedirs(IMAGE_DIR, exist_ok=True)

POSTED_COLUMNS = [
    "Category",
    "Name",
    "Picture",
    "Affiliate Link",
    "Description",
    "Status"
]
published = False


def start_browser():
    options = webdriver.ChromeOptions()

    # Existing Chrome Profile
    options.add_argument(
        r"--user-data-dir=D:\AutomationProfile"
    )

    # Optional: if profile name is Default
    options.add_argument("--profile-directory=Default")

    driver = webdriver.Chrome(options=options)

    driver.maximize_window()

    return driver


def post_pin(driver, image_path,title, description, link):

    try:

        wait = WebDriverWait(driver, 60)

        driver.get(
            "https://in.pinterest.com/pin-creation-tool/"
        )

        file_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "storyboard-upload-input"
                )
            )
        )

        file_input.send_keys(
            os.path.abspath(image_path)
        )

        print("Image selected")

        # Wait until upload box disappears
        wait.until_not(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "storyboard-upload-input"
                )
            )
        )

        print("Image uploaded successfully")
        title = remove_non_bmp(title)
        description = remove_non_bmp(description)
        # -------------------------
        # TITLE4o
        
        # -------------------------
        title_box = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "storyboard-selector-title")
            )
        )

        title_box.clear()
        title_box.send_keys(title)

        print("Title filled")

        # -------------------------
        # DESCRIPTION
        # -------------------------
        desc_editor = wait.until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    ".public-DraftEditor-content"
                )
            )
        )

        driver.execute_script(
            """
            arguments[0].focus();
            """,
            desc_editor
        )

        desc_editor.send_keys(description[:800])

        print("Description filled")

        # -------------------------
        # LINK
        # -------------------------
        link_box = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "WebsiteField")
            )
        )

        link_box.clear()
        link_box.send_keys(link)

        print("Link filled")

        return True

    except Exception as e:

        print("Upload failed")
        print(e)

        return False
    

def publish_pin(driver):

    try:

        wait = WebDriverWait(driver, 120)

        # Wait until Pinterest finishes saving
        wait.until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    '[data-test-id="saving-status-saved"]'
                )
            )
        )

        print("Changes stored detected")

        publish_btn = wait.until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    '[data-test-id="storyboard-creation-nav-done"] button'
                )
            )
        )

        

        

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            publish_btn
        )

        time.sleep(2)

        from selenium.webdriver.common.action_chains import ActionChains

        ActionChains(driver)\
            .move_to_element(publish_btn)\
            .pause(1)\
            .click()\
            .perform()

        print("Publish clicked")

        

        print("Pin published")
        return True
            

        

    except Exception as e:

        print("Publish failed")
        print(e)

        return False

def remove_from_amazon_products(product_name):

    try:

        df = pd.read_excel(AMAZON_FILE)

        target = product_name.strip().lower()

        df = df[
            df["Name"]
            .astype(str)
            .str.strip()
            .str.lower()
            != target
        ]

        df.to_excel(
            AMAZON_FILE,
            index=False
        )

        print(
            f"Removed: {product_name}"
        )

    except Exception as e:

        print(e)
# =========================
# CREATE POSTED FILE
# =========================

def create_posted_file():
    if not os.path.exists(POSTED_FILE):
        pd.DataFrame(columns=POSTED_COLUMNS).to_excel(
            POSTED_FILE,
            index=False
        )
        print("Created posted_pins.xlsx")


# =========================
# LOAD POSTED PRODUCTS
# =========================

def get_posted_product_names():

    create_posted_file()

    df = pd.read_excel(POSTED_FILE)

    if "Name" not in df.columns:
        return set()

    return set(
        df["Name"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
    )

def remove_non_bmp(text):
    if text is None:
        return ""

    return "".join(
        ch for ch in str(text)
        if ord(ch) <= 0xFFFF
    )
# =========================
# DOWNLOAD IMAGE
# =========================

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
        )

        safe_name = safe_name[:100]

        ext = ".jpg"

        if ".png" in url.lower():
            ext = ".png"

        image_path = os.path.join(
            IMAGE_DIR,
            safe_name + ext
        )

        if os.path.exists(image_path):
            return image_path

        response = requests.get(
            url,
            timeout=30,
            stream=True
        )

        response.raise_for_status()

        with open(image_path, "wb") as f:
            for chunk in response.iter_content(1024):
                f.write(chunk)

        print(f"Downloaded image: {product_name}")

        return image_path

    except Exception as e:

        print(f"Image download failed: {product_name}")
        print(e)

        return None


# =========================
# PREPARE PINS
# =========================

def prepare_pins(limit):

    create_posted_file()

    posted_names = get_posted_product_names()

    amazon_df = pd.read_excel(AMAZON_FILE)

    pins_to_post = []

    for _, row in amazon_df.iterrows():

        product_name = str(row["Name"]).strip()

        if product_name.lower() in posted_names:
            continue
        if pd.isna(row["Picture"]):
            continue

        if pd.isna(row["Affiliate Link"]):
            continue
        image_path = download_image(
            row["Picture"],
            product_name
        )

        if not image_path:
            continue

        pin_data = {
            "Category": row["Category"],
            "Name": product_name,
            "Picture": row["Picture"],
            "Affiliate Link": row["Affiliate Link"],
            "Description": (
                ""
                if pd.isna(row["Description"])
                else str(row["Description"])
            ),
            "Image Path": image_path
        }

        pins_to_post.append(pin_data)

        if len(pins_to_post) >= limit:
            break

    return pins_to_post


# =========================
# SAVE POSTED PRODUCT
# =========================

def mark_as_posted(product):

    df = pd.read_excel(POSTED_FILE)

    new_row = {
        "Category": product["Category"],
        "Name": product["Name"],
        "Picture": product["Picture"],
        "Affiliate Link": product["Affiliate Link"],
        "Description": product["Description"],
        "Status": "Posted"
    }

    df = pd.concat(
        [df, pd.DataFrame([new_row])],
        ignore_index=True
    )

    df.to_excel(
        POSTED_FILE,
        index=False
    )

    print(
        f"Added to posted_pins.xlsx: {product['Name']}"
    )


# =========================
# TEST
# =========================

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            '"' + sys.argv[0] + '"',
            None,
            1
        )
        sys.exit()
    print("Reading current timeout values...")

    # Save original values
    original_ac, original_dc = get_current_timeouts()

    print_timeouts("Original Display Timeout")

    # Maximum timeout supported by Windows
    MAX_TIMEOUT = 0xFFFFFFFF

    print("\nSetting timeout to maximum...")

    set_display_timeout(MAX_TIMEOUT, MAX_TIMEOUT)

    print_timeouts("After Setting Maximum Timeout")
    number_of_pins = int(
        input("How many pins to prepare? ")
    )

    pins = prepare_pins(number_of_pins)

    print(f"\nPins Ready: {len(pins)}\n")

    driver = start_browser()

    try:

        for i, pin in enumerate(pins, start=1):
            
            print(f"\nPosting {i}/{len(pins)}")
            print(pin["Name"])
            product_name = pin["Name"]

            # Remove Amazon feature list
            product_name = product_name.split("|")[0].strip()

            # If still too long, cut at first comma
            if len(product_name) > 80 and "," in product_name:
                product_name = product_name.split(",")[0].strip()

            # Final length limit
            if len(product_name) > 80:
                product_name = product_name[:80].rsplit(" ", 1)[0]
            description = str(
                pin.get("Description", "")
            )

            if description.lower() == "nan":
                description = ""

            description = description[:800]
                
            

            success = post_pin(
                driver,
                pin["Image Path"],
                product_name,
                description,
                pin["Affiliate Link"]
            )
            wait = WebDriverWait(driver, 60)
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        '[data-test-id="saving-status-saved"]'
                    )
                )
            )
            if success:

                    print("\nDraft created successfully")
                    print("Please click Publish manually.")

                    publish_pin(driver)
                    publish_pin(driver)
                    mark_as_posted(pin)
                    remove_from_amazon_products(pin["Name"])
            
                
            if published:
                mark_as_posted(pin)
                remove_from_amazon_products(pin["Name"])
            if i==number_of_pins:
                break
            time.sleep(60)
            print("=========================================")
    except Exception as e:

        print("An error occurred during posting:")
        print(e)
    finally:
        driver.quit()
    print("Restoring original timeout values...")

    set_display_timeout(original_ac, original_dc)

    print_timeouts("Restored Original Timeout")

    print("\nDone!")