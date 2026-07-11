# 📌 Amazon to Pinterest Automation

Automate your affiliate marketing workflow by scraping Amazon products and publishing them to Pinterest automatically using Selenium.

The project now includes several improvements such as configurable scraping limits, duplicate prevention, automatic display timeout management during long-running sessions, and a robust queue-based workflow.

---

# 🚀 Features

## 🛒 Amazon Product Scraper (`amazon.py`)

### Product Collection

* ✅ Search Amazon using custom keywords
* ✅ Scrape multiple pages automatically
* ✅ Scrape a user-defined number of products
* ✅ Skip duplicate products already present in the Excel queue
* ✅ Continue scraping until the requested number of products is collected

### Product Information

Each product includes:

* Product Category
* Product Name
* Product Image URL
* Amazon Affiliate Link
* Product Description

### Excel Queue

Automatically stores products in:

```
data/amazon_products.xlsx
```

The queue prevents duplicates and allows products to be posted later.

### Automatic Display Timeout Management

Long scraping sessions no longer stop because the monitor turns off.

The scraper automatically:

* Reads the current Windows display timeout
* Saves the original AC/Battery timeout values
* Temporarily increases the display timeout while scraping
* Restores the user's original settings when the script finishes

Administrator privileges are requested automatically when required.

---

## 📌 Pinterest Automation (`pins.py`)

The Pinterest automation reads products directly from the Excel queue.

Features include:

* ✅ Read queued products automatically
* ✅ Download product images
* ✅ Upload images to Pinterest
* ✅ Fill Pin Title
* ✅ Fill Pin Description
* ✅ Add Affiliate Link
* ✅ Publish Pins automatically
* ✅ Archive successful posts
* ✅ Remove posted products from the queue
* ✅ Resume automatically from remaining products
* ✅ Prevent duplicate posting

---

# 🔄 Complete Workflow

```text
Search Keyword
      │
      ▼
 amazon.py
      │
      ▼
Scrape Amazon Products
      │
      ▼
Duplicate Check
      │
      ▼
amazon_products.xlsx
   (Pending Queue)
      │
      ▼
      pins.py
      │
      ▼
Read Next Product
      │
      ▼
Download Product Image
      │
      ▼
Create Pinterest Pin
      │
      ▼
Publish Pin
      │
      ├────────────► posted_pins.xlsx
      │              (Archive)
      │
      ▼
Remove Product From Queue
```

---

# 📂 Project Structure

```text
Amazon-Pinterest-Automation/
│
├── amazon.py
├── pins.py
│
├── data/
│   ├── amazon_products.xlsx
│   └── posted_pins.xlsx
│
├── images/
│
├── requirements.txt
└── README.md
```

---

# 📊 Excel Files

## `amazon_products.xlsx`

Stores products waiting to be published.

| Category | Name | Picture | Affiliate Link | Description |
| -------- | ---- | ------- | -------------- | ----------- |

---

## `posted_pins.xlsx`

Stores products that have already been published.

| Category | Name | Picture | Affiliate Link | Description | Posted Date |
| -------- | ---- | ------- | -------------- | ----------- | ----------- |

This archive prevents duplicate Pinterest posts.

---

# ⚙️ Requirements

* Python 3.10+
* Google Chrome
* Pinterest Account
* Amazon Associates Account
* Internet Connection
* Windows (Administrator privileges required for automatic display timeout management)

---

# 📦 Installation

Clone the repository:

```bash
git clone https://github.com/Revanth1310/amazon_to_pintrest.git

cd amazon_to_pintrest
```

Create a virtual environment:

```bash
python -m venv venv
```

Activate it.

### Windows

```bash
venv\Scripts\activate
```

### Linux / macOS

```bash
source venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

# 🛒 Running Amazon Scraper

Run:

```bash
python amazon.py
```
enter product names

The script will ask for:

```
Enter Search Query:
```

Example:

```
wireless earbuds
```

Then:

```
Enter No of Items to Scrape:
```

Example:

```
100
```

The scraper will:

1. Request Administrator permission (if needed)
2. Save current display timeout
3. Prevent the display from turning off
4. Search Amazon
5. Collect products
6. Skip duplicates
7. Save products into the queue
8. Restore the original display timeout automatically

---

# 📌 Running Pinterest Automation

Make sure your Chrome profile is already logged into Pinterest.

Configure your profile path:

```python
PROFILE_PATH = "<Crome Driver Path>"
```

Run:

```bash
python pins.py
```
enter no of pins to post

The automation will:

1. Read the next queued product
2. Download its image
3. Create a Pinterest Pin
4. Upload the image
5. Add title
6. Add description
7. Add affiliate link
8. Publish the pin
9. Archive the product
10. Remove it from the pending queue

---

# 📋 Queue Management

### Before Posting

```text
amazon_products.xlsx

Product A
Product B
Product C
```

After Product A is posted:

```text
amazon_products.xlsx

Product B
Product C
```

```text
posted_pins.xlsx

Product A
```

This guarantees every product is posted only once.

---

# 🛠 Technologies Used

* Python
* Selenium
* Pandas
* OpenPyXL
* Pyperclip
* Requests
* ChromeDriver Manager
* Windows PowerCfg
* Regular Expressions (Regex)

---

# 🔒 Safety Features

* Duplicate product detection
* Duplicate Pinterest prevention
* Automatic queue management
* Automatic display timeout restoration
* Administrator privilege detection
* Exception handling during scraping
* Resume from remaining products

---

# ⚠️ Notes

* Pinterest UI updates may require Selenium selector changes.
* Amazon page updates may require scraper adjustments.
* Respect Amazon Associates Program policies.
* Respect Pinterest automation policies.
* Keep Chrome and ChromeDriver up to date.
* Avoid excessive request rates to reduce the likelihood of rate limiting.

---

# 🔮 Planned Features

* Pinterest Board Selection
* AI-generated Pin Titles
* AI-generated Descriptions
* Product Price Filtering
* Rating-based Product Filtering
* Scheduled Pinterest Posting
* Multiple Pinterest Accounts
* Multiple Amazon Marketplaces
* Cloud Deployment
* GUI Version
* Analytics Dashboard
* Proxy Support

---

# 🤝 Contributing

Contributions, issues, and feature requests are welcome.

Feel free to fork the repository and submit a pull request.

---

# 📜 License

MIT License

---

# ⭐ Support

If you find this project useful, consider giving it a ⭐ on GitHub.

Your support helps improve the project and encourages future development.

Happy Automating! 🚀
