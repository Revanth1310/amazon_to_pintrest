import os
import random
import time
import json
import pandas as pd

from openpyxl import Workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from urllib.parse import urljoin
# ==========================================================
# CONFIG
# ==========================================================

FOLDER_PATH = "data"
FILE_NAME = "amazon_products.xlsx"
FILE_PATH = os.path.join(FOLDER_PATH, FILE_NAME)

AUTH_FILE = "auth.json"

SEARCH_WAIT = 2
PRODUCT_WAIT = 3
AFFILIATE_WAIT = 3

# Number of NEW products to collect in one run
PRODUCT_LIMIT = 20

AMAZON_URL = "https://www.amazon.in"

# Set HEADLESS = False for local testing.
# GitHub Actions should normally use True unless you configure Xvfb.
HEADLESS = True


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

        if "Name" not in df.columns:
            return set()

        return set(
            df["Name"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
        )

    except Exception:
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
    except Exception:
        df = pd.DataFrame(
            columns=[
                "Category",
                "Name",
                "Picture",
                "Affiliate Link",
                "Description"
            ]
        )

    # Extra duplicate protection
    if "Name" in df.columns:
        existing_names = set(
            df["Name"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
        )

        if str(name).strip().lower() in existing_names:
            print("Already exists in Excel -> Skip save")
            return False

    df.loc[len(df)] = [
        category,
        name,
        picture,
        affiliate_link,
        description
    ]

    df.to_excel(FILE_PATH, index=False)

    print("Saved To Excel")
    return True


# ==========================================================
# BROWSER
# ==========================================================

def launch_browser():
    """
    Launch Chromium with Playwright authentication state.

    auth.json must be a Playwright storage-state file created with:
        context.storage_state(path="auth.json")
    """

    if not os.path.exists(AUTH_FILE):
        raise FileNotFoundError(
            f"Authentication file not found: {AUTH_FILE}"
        )

    playwright = sync_playwright().start()

    browser = playwright.chromium.launch(
        headless=False,
    )

    context = browser.new_context(
        storage_state=AUTH_FILE,
        viewport={"width": 1920, "height": 1080}
    )

    page = context.new_page()

    print("Playwright browser started")
    print(f"Authentication state loaded: {AUTH_FILE}")

    return playwright, browser, context, page


# ==========================================================
# AMAZON SEARCH
# ==========================================================

def search_amazon(page, query):
    print(f"\nSearching Amazon for: {query}")

    page.goto(
        AMAZON_URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    search_box = page.locator("#twotabsearchtextbox")

    search_box.wait_for(
        state="visible",
        timeout=30000
    )

    search_box.fill(query)
    search_box.press("Enter")

    page.wait_for_url(
        "**/s?**",
        timeout=30000
    )

    print("\nSearch URL:")
    print(page.url)

    time.sleep(SEARCH_WAIT)


# ==========================================================
# PRODUCT LIST
# ==========================================================

def get_products(page):
    products = page.locator(
        'div.s-result-item[data-component-type="s-search-result"]'
    )

    valid_products = []

    count = products.count()

    for index in range(count):
        product = products.nth(index)

        try:
            asin = product.get_attribute("data-asin")

            if asin:
                valid_products.append(product)

        except Exception:
            continue

    return valid_products


# ==========================================================
# PRODUCT DETAILS
# ==========================================================

def get_product_name(page):
    try:
        title = page.locator("#productTitle").first

        title.wait_for(
            state="visible",
            timeout=15000
        )

        return title.inner_text().strip()

    except Exception:
        return ""


def get_product_image(page):
    try:
        img = page.locator("#landingImage").first

        img.wait_for(
            state="visible",
            timeout=15000
        )

        src = img.get_attribute("src")

        if src:
            return src

        # Fallback for lazy-loaded images
        src = img.get_attribute("data-old-hires")

        return src or ""

    except Exception:
        return ""


def get_product_description(page):
    try:
        bullets = page.locator(
            "#feature-bullets li span.a-list-item"
        )

        desc = []

        for index in range(bullets.count()):
            try:
                text = bullets.nth(index).inner_text().strip()

                if text:
                    desc.append(text)

            except Exception:
                continue

        return " | ".join(desc)

    except Exception:
        return ""


def get_product_category(page):
    try:
        category = page.locator(
            "#amzn-ss-category-content"
        ).first

        if category.count() > 0:
            return category.inner_text().strip()

    except Exception:
        pass

    return ""


# ==========================================================
# AFFILIATE LINK
# ==========================================================

def get_affiliate_link(page):
    """
    Uses the existing Amazon Associates SiteStripe selectors
    from the Selenium version.

    If the affiliate controls are not available, the current
    product URL is returned as a fallback.
    """

    current_url = page.url

    try:
        time.sleep(2)

        get_link_btn = page.locator(
            "#amzn-ss-get-link-button"
        ).first

        get_link_btn.wait_for(
            state="visible",
            timeout=15000
        )

        get_link_btn.click()

        print("Clicked Get Link")

        time.sleep(AFFILIATE_WAIT)

        copy_btn = page.locator(
            "#amzn-ss-copy-affiliate-link-btn-announce"
        ).first

        copy_btn.wait_for(
            state="visible",
            timeout=15000
        )

        # Click copy button if available.
        # We don't depend on the system clipboard because
        # clipboard access is unreliable in GitHub Actions.
        copy_btn.click()

        print("Clicked Copy Affiliate Link")

        time.sleep(2)

        # Try common input/text fields created by SiteStripe.
        selectors = [
            "#amzn-ss-text-shortlink",
            "#amzn-ss-text-affiliate-link",
            "input[id*='affiliate-link']",
            "input[id*='shortlink']",
            "textarea"
        ]

        for selector in selectors:
            try:
                loc = page.locator(selector).first

                if loc.count() > 0:
                    value = loc.input_value(timeout=3000)

                    if value and value.startswith("http"):
                        print("Affiliate Link Retrieved")
                        return value.strip()

            except Exception:
                continue

        # Try reading text from the SiteStripe area.
        try:
            site_stripe = page.locator(
                "[id*='amzn-ss']"
            )

            text = site_stripe.inner_text(timeout=3000)

            for line in text.splitlines():
                line = line.strip()

                if line.startswith("http"):
                    print("Affiliate Link Retrieved From Text")
                    return line

        except Exception:
            pass

        print("Could not read generated affiliate link.")
        print("Using current product URL as fallback.")

        return current_url

    except Exception as e:
        print("Affiliate Link Error:", e)
        print("Using current product URL as fallback.")

        return current_url


# ==========================================================
# SCRAPE PRODUCT
# ==========================================================

def scrape_product(page):
    category = get_product_category(page)

    name = get_product_name(page)

    image = get_product_image(page)

    description = get_product_description(page)

    affiliate_link = get_affiliate_link(page)

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
    page,
    page_number,
    existing_products
):
    global i, c

    products = get_products(page)

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
            # Re-read products because the page can change
            # after returning from a product page.
            products = get_products(page)

            if index >= len(products):
                break

            product = products[index]

            link_element = product.locator(
                "a.a-link-normal.s-no-outline"
            ).first

            link_element.wait_for(
                state="attached",
                timeout=10000
            )

            product = products[index]

            asin = product.get_attribute("data-asin")

            if not asin:
                print("ASIN Empty -> Skip")
                continue

            product_url = f"https://www.amazon.in/dp/{asin}"

            print(f"\nOpening Product {index + 1}")
            print(f"ASIN: {asin}")
            print(f"Product URL: {product_url}")

            product_page = page.context.new_page()

            try:
                product_page.goto(
                    product_url,
                    wait_until="domcontentloaded",
                    timeout=60000
                )

                time.sleep(PRODUCT_WAIT)

                data = scrape_product(product_page)

            finally:
                product_page.close()

            product_name = (
                data["name"]
                .strip()
                .lower()
            )

            if not product_name:
                print("Product Name Empty")

                continue

            if product_name in existing_products:
                print("Already Exists -> Skip")

                continue

            saved = save_product(
                data["category"],
                data["name"],
                data["image"],
                data["affiliate_link"],
                data["description"]
            )

            if saved:
                existing_products.add(
                    product_name
                )

                print(
                    f"Saved: {data['name']}"
                )

                i += 1

                print(
                    f"Scraped {i}/{c}"
                )

            time.sleep(1)

        except Exception as e:
            print(
                f"Error Product {index + 1}: {e}"
            )

    return True


# ==========================================================
# NEXT PAGE
# ==========================================================

def goto_next_page(page):
    try:
        next_btn = page.locator(
            "a.s-pagination-next"
        ).first

        if next_btn.count() == 0:
            return False

        if not next_btn.is_visible():
            return False

        disabled = next_btn.get_attribute("aria-disabled")

        if disabled == "true":
            return False

        next_btn.scroll_into_view_if_needed()

        time.sleep(1)

        next_btn.click()

        # Wait for the next page/search result content.
        page.wait_for_load_state(
            "domcontentloaded",
            timeout=30000
        )

        time.sleep(3)

        return True

    except Exception:
        return False


# ==========================================================
# SCRAPE ALL PAGES
# ==========================================================

def scrape_all_pages(
    page,
    existing_products
):
    page_number = 1

    while True:

        continue_scraping = process_current_page(
            page,
            page_number,
            existing_products
        )

        if not continue_scraping:
            print(
                f"\nReached requested limit ({c} products)."
            )
            break

        moved = goto_next_page(page)

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

    queries = [
        "women's trendy dresses",
        "women's casual tops",
        "women's ethnic wear",
        "women's handbags",
        "men's casual shirts",
        "men's t shirts",
        "men's ethnic wear",
        "kitchen storage organizers",
        "kitchen gadgets and tools",
        "non stick cookware set"
    ]

    query = random.choice(queries)

    print("=" * 60)
    print("Amazon Product Automation - Playwright")
    print("=" * 60)

    print(f"Random Query: {query}")

    global c
    c = PRODUCT_LIMIT

    global i
    i = 0

    create_excel_if_not_exists()

    existing_products = load_existing_products()

    print(
        f"Loaded {len(existing_products)} Existing Products"
    )

    playwright = None
    browser = None
    context = None
    page = None

    try:
        playwright, browser, context, page = launch_browser()

        search_amazon(
            page,
            query
        )

        scrape_all_pages(
            page,
            existing_products
        )

        print(
            "\nScraping Completed"
        )

        print(
            f"New Products Saved: {i}"
        )

    except Exception as e:
        print(
            f"\nAutomation Error: {e}"
        )
        raise

    finally:
        if context:
            context.close()

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

        print("\nBrowser closed.")
        print("Done!")


if __name__ == "__main__":
    main()