import os
import time
import pyperclip
import pandas as pd

from openpyxl import Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
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

# ==========================================================
# CONFIG
# ==========================================================

PROFILE_PATH = r"D:\AutomationProfile"

FOLDER_PATH = "data"
FILE_NAME = "amazon_products.xlsx"
FILE_PATH = os.path.join(FOLDER_PATH, FILE_NAME)

SEARCH_WAIT = 2
PRODUCT_WAIT = 3
AFFILIATE_WAIT = 3


# ==========================================================
# EXCEL FUNCTIONS
# ==========================================================

def create_excel_if_not_exists():

    os.makedirs(FOLDER_PATH, exist_ok=True)

    if not os.path.exists(FILE_PATH):

        wb = Workbook()
        ws = wb.active
        ws.title = "Products"

        ws.append([
            "Category",
            "Name",
            "Picture",
            "Affiliate Link",
            "Description"
        ])

        wb.save(FILE_PATH)

        print(f"Created Excel: {FILE_PATH}")

    else:
        print(f"Excel Exists: {FILE_PATH}")


def load_existing_products():

    try:

        df = pd.read_excel(FILE_PATH)

        return set(
            df["Name"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

    except:

        return set()


def save_product(
    category,
    name,
    picture,
    affiliate_link,
    description
):

    try:
        df = pd.read_excel(FILE_PATH)

    except:
        df = pd.DataFrame(
            columns=[
                "Category",
                "Name",
                "Picture",
                "Affiliate Link",
                "Description"
            ]
        )

    df.loc[len(df)] = [
        category,
        name,
        picture,
        affiliate_link,
        description
    ]

    df.to_excel(FILE_PATH, index=False)

    print("Saved To Excel")


# ==========================================================
# DRIVER
# ==========================================================

def launch_browser():

    options = webdriver.ChromeOptions()

    options.add_argument(
        f"--user-data-dir={PROFILE_PATH}"
    )

    driver = webdriver.Chrome(
        service=Service(
            ChromeDriverManager().install()
        ),
        options=options
    )

    driver.maximize_window()

    return driver


# ==========================================================
# AMAZON SEARCH
# ==========================================================

def search_amazon(driver, query):

    driver.get("https://www.amazon.in")

    search_box = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.ID, "twotabsearchtextbox")
        )
    )

    search_box.clear()
    search_box.send_keys(query)
    search_box.send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(
        lambda d: "/s?" in d.current_url
    )

    print("\nSearch URL:")
    print(driver.current_url)

    time.sleep(SEARCH_WAIT)


# ==========================================================
# PRODUCT LIST
# ==========================================================

def get_products(driver):

    all_products = driver.find_elements(
        By.CSS_SELECTOR,
        'div.s-result-item[data-component-type="s-search-result"]'
    )

    valid_products = []

    for product in all_products:

        asin = product.get_attribute(
            "data-asin"
        )

        if asin:
            valid_products.append(product)

    return valid_products


# ==========================================================
# PRODUCT DETAILS
# ==========================================================

def get_product_name(driver):

    try:

        return WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.ID, "productTitle")
            )
        ).text.strip()

    except:
        return ""


def get_product_image(driver):

    try:

        img = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.ID, "landingImage")
            )
        )

        return img.get_attribute("src")

    except:
        return ""


def get_product_description(driver):

    try:

        bullets = driver.find_elements(
            By.CSS_SELECTOR,
            "#feature-bullets li span.a-list-item"
        )

        desc = []

        for bullet in bullets:

            text = bullet.text.strip()

            if text:
                desc.append(text)

        return " | ".join(desc)

    except:
        return ""


def get_product_category(driver):

    try:

        return driver.find_element(
            By.ID,
            "amzn-ss-category-content"
        ).text.strip()

    except:

        return ""


# ==========================================================
# AFFILIATE LINK
# ==========================================================

def get_affiliate_link(driver):

    try:

        time.sleep(2)

        # CLICK GET LINK
        get_link_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (
                    By.ID,
                    "amzn-ss-get-link-button"
                )
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            get_link_btn
        )

        print("Clicked Get Link")

        time.sleep(AFFILIATE_WAIT)

        # CLICK COPY AFFILIATE LINK
        copy_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (
                    By.ID,
                    "amzn-ss-copy-affiliate-link-btn-announce"
                )
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            copy_btn
        )

        print("Clicked Copy Affiliate Link")

        time.sleep(2)

        affiliate_link = pyperclip.paste()

        print("Affiliate Link Copied")

        try:

            close_btn = driver.find_element(
                By.CSS_SELECTOR,
                "button.a-button-close"
            )

            close_btn.click()

        except:
            pass

        return affiliate_link

    except Exception as e:

        print(
            "Affiliate Link Error:",
            e
        )

        return driver.current_url


# ==========================================================
# SCRAPE PRODUCT
# ==========================================================

def scrape_product(driver):

    category = get_product_category(driver)

    name = get_product_name(driver)

    image = get_product_image(driver)

    description = get_product_description(driver)

    affiliate_link = get_affiliate_link(driver)

    return {
        "category": category,
        "name": name,
        "image": image,
        "affiliate_link": affiliate_link,
        "description": description
    }


# ==========================================================
# PROCESS PAGE
# ==========================================================

def process_current_page(
    driver,
    page_number,
    existing_products,
):
    global i, c
    products = get_products(driver)

    print(
        f"\n========== PAGE {page_number} =========="
    )

    print(
        f"Products Found: {len(products)}"
    )
    
    
        
    for index in range(len(products)):
        if i >= c:
            return False
        try:

            products = get_products(driver)

            product = products[index]

            link_element = product.find_element(
                By.CSS_SELECTOR,
                "a.a-link-normal.s-no-outline"
            )

            product_url = link_element.get_attribute(
                "href"
            )

            print(
                f"\nOpening Product {index + 1}"
            )

            print(product_url)

            driver.execute_script(
                "window.open(arguments[0]);",
                product_url
            )

            driver.switch_to.window(
                driver.window_handles[-1]
            )

            time.sleep(PRODUCT_WAIT)

            data = scrape_product(driver)

            product_name = (
                data["name"]
                .strip()
                .lower()
            )

            if not product_name:

                print(
                    "Product Name Empty"
                )

            elif product_name in existing_products:

                print(
                    "Already Exists -> Skip"
                )

            else:

                save_product(
                    data["category"],
                    data["name"],
                    data["image"],
                    data["affiliate_link"],
                    data["description"]
                )

                existing_products.add(
                    product_name
                )

                print(
                    f"Saved: {data['name']}"
                )
                i += 1
                print(f"Scraped {i}/{c}")

                

            driver.close()

            driver.switch_to.window(
                driver.window_handles[0]
            )

            time.sleep(1)

        except Exception as e:

            print(
                f"Error Product {index + 1}: {e}"
            )

            try:

                if len(
                    driver.window_handles
                ) > 1:

                    driver.close()

                    driver.switch_to.window(
                        driver.window_handles[0]
                    )

            except:
                pass
    return True


# ==========================================================
# NEXT PAGE
# ==========================================================

def goto_next_page(driver):

    try:

        next_btn = driver.find_element(
            By.CSS_SELECTOR,
            "a.s-pagination-next"
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            next_btn
        )

        time.sleep(2)

        driver.execute_script(
            "arguments[0].click();",
            next_btn
        )

        time.sleep(3)

        return True

    except:

        return False


# ==========================================================
# SCRAPE ALL PAGES
# ==========================================================

def scrape_all_pages(
    driver,
    existing_products,
    
):

    page_number = 1

    while True:

        continue_scraping = process_current_page(
            driver,
            page_number,
            existing_products,
        )

        if not continue_scraping:
            print(f"\nReached requested limit ({c} products).")
            break
        moved = goto_next_page(driver)

        if not moved:

            print(
                "\nNo More Pages Found"
            )

            break

        page_number += 1

        print(
            f"\nMoving To Page {page_number}"
        )


# ==========================================================
# MAIN
# ==========================================================

def main():
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
    query = input(
        "Enter Search Query: "
    )
    global c
    c = int(input("Enter No of Items to Scrape: "))
    global i
    i = 0
    create_excel_if_not_exists()

    existing_products = (
        load_existing_products()
    )

    print(
        f"Loaded {len(existing_products)} Existing Products"
    )
    

    driver = launch_browser()

    try:

        search_amazon(
            driver,
            query
        )

        scrape_all_pages(
            driver,
            existing_products,
            
        )

        print(
            "\nScraping Completed"
        )

        input(
            "\nPress Enter To Close..."
        )

    finally:

        driver.quit()
    print("Restoring original timeout values...")

    set_display_timeout(original_ac, original_dc)

    print_timeouts("Restored Original Timeout")

    print("\nDone!")

if __name__ == "__main__":
    
    main()