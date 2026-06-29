# 🚗 Car Scraper Project

A Python-based web scraping tool that collects used car listings from **4 major Malaysian car platforms** — all in one project.

I built this to automate the tedious process of manually browsing through thousands of car listings across different websites. Instead of opening each site one by one and scrolling through pages, this project does all the heavy lifting for you and neatly exports everything into organized Excel files.

---

## 🌐 Platforms Covered

| Platform | Website | Scraping Method | Speed |
|----------|---------|-----------------|-------|
| **Mudah** | [mudah.my](https://www.mudah.my) | API + Detail Page Scraping | ~2-3 min per page |
| **Carsome** | [carsome.my](https://www.carsome.my) | REST API | ~5 sec per page |
| **Carro** | [carro.co](https://carro.co) | REST API | ~5 sec per page |
| **Carlist** | [carlist.my](https://www.carlist.my) | Browser Automation (Playwright) | ~3-5 min per page |

> **Why the speed difference?** Mudah and Carlist require visiting each car's individual detail page to get full specs (a 2-step process), while Carsome and Carro expose all the data through their APIs in a single request.

---

## 📊 What Data Does It Collect?

Each scraper exports a **timestamped Excel file** (e.g., `Mudah-2026-06-29-10-00-00.xlsx`) with the following **core columns**:

| Column | Description |
|--------|-------------|
| Extraction Time | Exact timestamp when each car was scraped |
| Title | Car name (e.g., Honda City V 1.5) |
| Brand | Manufacturer (e.g., Honda, Toyota, Perodua) |
| Model | Specific model name |
| Variant | Trim level or variant |
| Year | Manufacturing year |
| Price | Listing price in RM |
| Mileage | Odometer reading |
| Engine CC | Engine capacity |
| Transmission | Auto / Manual |
| State | State location (e.g., Selangor, Johor) |
| Area | City or area (e.g., Petaling Jaya, Kuching) |
| Listing Link | Direct URL to the listing |
| Main Image URL | Link to the car's main photo |

> On top of these core columns, each platform also exports its own **unique bonus columns** — things like warranty status, seller type, loan details, interior/exterior photo counts, and more.

---

## ⚙️ How It Works

Each scraper has a simple **configuration block** at the top of the file where you can toggle between two modes:

```python
# --- OPTION A: TEST MODE (Quick run, just a few pages) ---
PAGES_TO_SCRAPE = 2

# --- OPTION B: FULL EXTRACTION MODE (Scrapes everything, auto-stops when done) ---
# PAGES_TO_SCRAPE = 99999
```

- **Option A** is great for quick demos or testing — it grabs just a small sample.
- **Option B** will keep scraping until there are no more cars left on the site, then stops automatically.

To switch modes, simply comment one line and uncomment the other.

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.8+** installed on your machine
- A terminal (Mac Terminal, Windows Command Prompt, or PowerShell)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/Alimieee/Car_Scraper_Project.git
   cd Car_Scraper_Project
   ```

2. **Create and activate a virtual environment**

   **Mac / Linux:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

   **Windows:**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install Playwright browsers** (required for Carlist)
   ```bash
   playwright install
   ```

### Running the Scrapers

Open a terminal, make sure your virtual environment is activated, then run any of the four scrapers:

```bash
# Mudah
cd mudah_site
python3 mudah_deep_scraper.py

# Carsome
cd carsome_site
python3 carsome_deep_scraper.py

# Carro
cd carro_site
python3 carro_deep_scraper.py

# Carlist
cd carlist_site
python3 carlist_deep_scraper.py
```

> 💡 **Tip:** You can run all four scrapers simultaneously by opening 4 separate terminal windows!

---

## 📁 Project Structure

```
Car_Scraper_Project/
├── mudah_site/
│   └── mudah_deep_scraper.py       # Mudah.my scraper
├── carsome_site/
│   └── carsome_deep_scraper.py     # Carsome.my scraper
├── carro_site/
│   └── carro_deep_scraper.py       # Carro.co scraper
├── carlist_site/
│   └── carlist_deep_scraper.py     # Carlist.my scraper
├── requirements.txt                # Python dependencies
├── README_FOR_BOSS.txt             # Quick-start guide
└── .gitignore
```

---

## 🛠 Tech Stack

| Tool | Purpose |
|------|---------|
| **Python 3** | Core programming language |
| **Requests** | Sending HTTP requests to APIs |
| **BeautifulSoup4** | Parsing HTML content |
| **Playwright** | Browser automation for Cloudflare-protected sites |
| **Pandas** | Data processing and structuring |
| **OpenPyXL** | Exporting data to formatted Excel files |

---

## 📝 Notes

- **Carlist** uses Playwright (a headless browser) because the site is protected by Cloudflare. You'll see a Chrome window open and navigate automatically — that's normal! Don't close it.
- Each run generates a **new Excel file** with a timestamp in the filename, so your previous data is never overwritten.
- The scrapers include built-in **delays between requests** to be respectful to the servers and avoid getting blocked.
- **Auto-stop** is built in — when set to full extraction mode, the scraper will automatically stop when it detects there are no more cars to collect.

---

## 👤 Author

**Alimi** — Built as a data collection tool for the Malaysian used car market.

---
