# 📌 Amazon to Pinterest Automation

Automate your affiliate marketing workflow by scraping products from Amazon and publishing them to Pinterest automatically.

This project consists of two Python automation scripts:

* **amazon.py** – Scrapes products from Amazon and stores them in an Excel queue.
* **pins.py** – Reads products from the queue, creates Pinterest Pins automatically, removes posted products from the queue, and archives them in a separate file.

---

# 🚀 Features

## Amazon Product Scraper (`amazon.py`)

✅ Search Amazon using product keywords

✅ Scrape product information

* Product Name
* Product Image URL
* Product Link(Affliate Link)
* Product Description
* Category

✅ Store products in Excel

✅ Bulk product collection

✅ Queue-based workflow

---

## Pinterest Automation (`pins.py`)

✅ Read products from `amazon_products.xlsx`

✅ Download product images automatically

✅ Upload images to Pinterest

✅ Fill Pin Title

✅ Fill Pin Description

✅ Add Destination Link

✅ Publish Pins automatically

✅ Remove successfully posted products from queue

✅ Archive posted products in `posted_pins.xlsx`

✅ Prevent duplicate posting

✅ Resume processing from remaining products

---

# 🔄 Workflow

```text
Product Keywords
        │
        ▼
     amazon.py
        │
        ▼
 amazon_products.xlsx
   (Pending Queue)
        │
        ▼
      pins.py
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
        │                (Posted Archive)
        │
        ▼
 Remove Product From
 amazon_products.xlsx
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

## amazon_products.xlsx

Stores products waiting to be posted.

| Category    | Name        | Picture   | Affiliate Link | Description         |
| ----------- | ----------- | --------- | -------------- | ------------------- |
| Electronics | Smart Watch | Image URL | Product URL    | Product Description |

---

## posted_pins.xlsx

Stores products that have already been posted successfully.

| Category | Name | Picture | Affiliate Link | Description | Posted Date |
| -------- | ---- | ------- | -------------- | ----------- | ----------- |

This file acts as an archive and helps prevent duplicate postings.

---

# ⚙️ Requirements

* Python 3.10+
* Google Chrome
* Pinterest Account
* Internet Connection

---

# 📦 Installation

Clone the repository:

```bash
git clone https://github.com/yourusername/amazon-pinterest-automation.git

cd amazon-pinterest-automation
```

Create a virtual environment:

```bash
python -m venv venv
```

Activate it:

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

# 🔍 Running Amazon Scraper

Edit product keywords inside `amazon.py`.

Run:

```bash
python amazon.py
```

The scraper will:

1. Search Amazon
2. Collect product details
3. Save products into:

```text
data/amazon_products.xlsx
```

---

# 📌 Running Pinterest Automation

Ensure your Chrome profile is already logged into Pinterest.

Configure your profile path:

```python
PROFILE_PATH = r"D:\AutomationProfile"
```

Run:

```bash
python pins.py
```

The bot will:

1. Read products from `amazon_products.xlsx`
2. Download product images
3. Create Pinterest Pins
4. Add title, description, and link
5. Publish the pin
6. Save product details to `posted_pins.xlsx`
7. Remove the product from `amazon_products.xlsx`

---

# 📋 Queue Management

### Before Posting

| Product   |
| --------- |
| Product A |
| Product B |
| Product C |

### After Product A is Posted

#### amazon_products.xlsx

| Product   |
| --------- |
| Product B |
| Product C |

#### posted_pins.xlsx

| Product   |
| --------- |
| Product A |

This ensures that products are never posted twice.

---

# 🛠 Technologies Used

* Python
* Selenium
* OpenPyXL
* Pandas
* Requests
* ChromeDriver Manager

---

# ⚠️ Important Notes

* Pinterest UI changes may require selector updates.
* Amazon page structure changes may require scraper updates.
* Use reasonable posting frequency.
* Ensure compliance with Amazon Affiliate and Pinterest policies.
* Keep Chrome and ChromeDriver updated.

---

# 🔮 Future Enhancements

* Pinterest Board Selection
* Scheduled Posting
* AI Generated Descriptions
* Multiple Pinterest Accounts
* Cloud Deployment
* Analytics Dashboard
* Automated Product Filtering

---

# 🤝 Contributing

Contributions, issues, and feature requests are welcome.

Feel free to fork the repository and submit a pull request.

---

# 📜 License

MIT License

---

# ⭐ Support

If this project helped you, please give it a ⭐ on GitHub.

Happy Automating! 🚀
