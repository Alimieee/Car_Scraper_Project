"""
Carsome.my Deep Car Scraper
============================
Extracts ALL available car data from Carsome's listing API.

Unlike Mudah.my where we needed to visit each car's detail page,
Carsome's API already returns 89 fields per car in the listing response —
including brand, model, variant, year, engine, mileage, transmission,
fuel type, color, price, monthly payment, location, VIN, license plate,
warranty status, features, images, and more.

We also construct the detail page URL for each car so you can visit it.

Usage: source ../mudah_site/venv/bin/activate && python3 carsome_deep_scraper.py
"""

import requests
import pandas as pd
import json
import time
import re
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from datetime import datetime


# ──────────────────────────────────────────────
# CONFIGURATION - CHOOSE ONE OPTION
# ──────────────────────────────────────────────

# --- OPTION A: TEST MODE (Quick run for your boss) ---
PAGES_TO_SCRAPE = 2

# --- OPTION B: FULL EXTRACTION MODE (Auto-stops when no cars left) ---
# PAGES_TO_SCRAPE = 99999
CARS_PER_PAGE = 18           # Carsome default is 18 per page
DELAY_BETWEEN_PAGES = 2      # Seconds between page requests
OUTPUT_FILE = f"Carsome-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/149.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
}

PAYLOAD = {"isShowPreList": True}

# Resilient session with retries
session = requests.Session()
retry = Retry(connect=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
adapter = HTTPAdapter(max_retries=retry)
session.mount("http://", adapter)
session.mount("https://", adapter)


# ──────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────

def build_detail_url(car):
    """Construct the Carsome detail page URL for a car."""
    brand = car.get("brandName", "").lower().replace(" ", "-")
    model = car.get("modelName", "").lower().replace(" ", "-")
    variant = car.get("carVariant", "").lower().replace(" ", "-")
    engine = str(car.get("carEngine", "")).replace(" ", "-")
    year = car.get("carYear", "")
    car_no = car.get("carNo", "").lower()

    slug = f"{year}-{brand}-{model}-{variant}-{engine}"
    # Clean up double dashes and trailing dashes
    slug = re.sub(r"-+", "-", slug).strip("-")

    return f"https://www.carsome.my/buy-car/{brand}/{model}/{slug}/{car_no}"


def format_price(price):
    """Format a numeric price to readable format."""
    if price and isinstance(price, (int, float)) and price > 0:
        return f"RM {price:,.0f}"
    return "N/A"


def extract_car_data(car):
    """Extract and organize all fields from a single car's listing data."""
    # Build the detail page URL
    detail_url = build_detail_url(car)

    # Pricing info
    all_price = car.get("allPrice", {})
    campaign_info = car.get("campaignInfo", {})
    loan_config = car.get("loanConfig", {})

    # Feature highlights
    highlights = car.get("highLightTags", [])
    highlight_str = " | ".join(highlights) if highlights else "N/A"

    # Car taggings (e.g., "Family Drive", "Budget Friendly")
    taggings = car.get("carTaggings", [])
    tagging_names = [t.get("tagName", "") for t in taggings if t.get("tagName")]
    tagging_str = " | ".join(tagging_names) if tagging_names else "N/A"

    # Campaign info
    campaigns = campaign_info.get("campaigns", [])
    campaign_names = [c.get("campaignName", "") for c in campaigns if c.get("campaignName")]
    campaign_str = " | ".join(campaign_names) if campaign_names else "N/A"

    # Images
    outer_images = car.get("imagesOuter", [])
    inner_images = car.get("imagesInner", [])
    main_image = car.get("image", "N/A")

    # Banner info
    banners = car.get("banners", [])
    banner_texts = []
    for b in banners:
        if isinstance(b, dict):
            text = b.get("text", b.get("title", ""))
            if text:
                banner_texts.append(text)
    banner_str = " | ".join(banner_texts) if banner_texts else "N/A"

    # Warranty status
    warranty = "Yes (1 Year)" if car.get("hasOneYearWarranty") == 1 else "No"

    # Condition mapping
    condition_map = {1: "Used", 2: "New", 3: "Recon"}
    condition = condition_map.get(car.get("carCondition"), str(car.get("carCondition", "N/A")))

    # Clean Title to remove year
    title = car.get("carName", "N/A")
    words = title.split()
    if words and words[0].isdigit() and len(words[0]) == 4:
        words.pop(0)
    clean_title = " ".join(words) if words else title

    return {
        # ── IDENTITY ──
        "Extraction Time": datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
        "Car ID": car.get("carId", car.get("id", "N/A")),
        "Stock No": car.get("carNo", "N/A"),
        "Car Name": clean_title,
        "Brand": car.get("brandName", "N/A"),
        "Model": car.get("modelName", "N/A"),
        "Variant": car.get("carVariant", "N/A"),
        "Year": car.get("carYear", "N/A"),
        "Car Age": car.get("carAge", "N/A"),

        # ── SPECS ──
        "Engine (L)": car.get("carEngine", "N/A"),
        "Transmission": car.get("carTransmissionName", "N/A"),
        "Fuel Type": car.get("fuelTypeName", "N/A"),
        "Color": car.get("colorName", "N/A"),
        "Seats": car.get("seat", "N/A"),
        "Mileage (km)": car.get("carMileage", "N/A"),
        "Condition": condition,

        # ── PRICING ──
        "Selling Price (RM)": car.get("price", "N/A"),
        "Original Price (RM)": all_price.get("originalPrice", car.get("expSellingPrice", "N/A")),
        "Discount (RM)": all_price.get("campaignDiscountAmount", "N/A"),
        "Monthly Payment (RM)": car.get("monthPayment", car.get("monthPrice", "N/A")),
        "Monthly (by Car Age)": all_price.get("monthPriceWithCarAge", "N/A"),

        # ── LOAN CONFIG ──
        "Down Payment Rate": f"{float(loan_config.get('downPaymentRate', 0)) * 100:.0f}%" if loan_config.get("downPaymentRate") else "N/A",
        "Loan Period (Years)": loan_config.get("loanPeriod", "N/A"),
        "Loan Rate (%)": loan_config.get("loanRate", "N/A"),
        "Loan Provider": loan_config.get("provider", "N/A"),

        # ── LOCATION ──
        "Location": car.get("location", "N/A"),
        "Store Name": car.get("storeName", "N/A"),
        "State": car.get("stateName", "N/A"),
        "Area": car.get("locationName", "N/A"),

        # ── REGISTRATION ──
        "License Plate": car.get("licensePlate", "N/A"),
        "VIN Code": car.get("vinCode", "N/A"),
        "Warranty": warranty,

        # ── CLASSIFICATION ──
        "Car Type": car.get("carTypeName", "N/A"),
        "Car Tag": car.get("carTag", "N/A"),
        "Category Tags": tagging_str,
        "Seller Type": car.get("dealerTypeName", "N/A"),
        "Availability": car.get("commodityStatusName", "N/A"),

        # ── FEATURES ──
        "Feature Highlights": highlight_str,
        "Active Promotions": campaign_str,
        "Banners": banner_str,

        # ── DATES ──
        "Listing Date": car.get("carListingDate", "N/A"),

        # ── IMAGES ──
        "Main Image URL": main_image,
        "Exterior Photos": len(outer_images),
        "Interior Photos": len(inner_images),
        "Total Photos": len(outer_images) + len(inner_images),

        # ── LINKS ──
        "Detail Page URL": detail_url,
    }


def scrape_page(page_num):
    """Scrape a single page from Carsome's listing API."""
    url = f"https://www.carsome.my/website/v2/car/list/{page_num}/{CARS_PER_PAGE}/1/"
    response = session.post(url, headers=HEADERS, json=PAYLOAD, timeout=15)

    if response.status_code != 200:
        print(f"  [!] Page {page_num} returned status {response.status_code}")
        return [], 0

    data = response.json()
    if data.get("code") not in (200, "0000"):
        print(f"  [!] API returned code {data.get('code')}: {data.get('msg')}")
        return [], 0

    api_data = data.get("data", {})
    car_list = api_data.get("carList", [])
    total_cars = api_data.get("total", 0)

    return car_list, total_cars


# ──────────────────────────────────────────────
# MAIN SCRAPING LOGIC
# ──────────────────────────────────────────────

def main():
    all_cars = []

    print("╔══════════════════════════════════════════════════════╗")
    print("║   Carsome.my Deep Car Scraper                      ║")
    print(f"║   Scraping {PAGES_TO_SCRAPE} page(s) × {CARS_PER_PAGE} cars/page                    ║")
    print("╚══════════════════════════════════════════════════════╝")

    total_available = 0

    for page in range(1, PAGES_TO_SCRAPE + 1):
        print(f"\n{'='*55}")
        print(f"  📋 Fetching Page {page}...")
        print(f"{'='*55}")

        car_list, total = scrape_page(page)
        if total > 0:
            total_available = total

        if not car_list:
            print(f"  [!] No cars found on page {page}. Auto-stopping pagination.")
            break

        print(f"  [+] Got {len(car_list)} cars (total available: {total_available})")

        for i, car in enumerate(car_list, 1):
            extracted = extract_car_data(car)
            all_cars.append(extracted)
            name = extracted.get("Car Name", "Unknown")
            price = extracted.get("Selling Price (RM)", "N/A")
            print(f"    [{i}/{len(car_list)}] ✅ {name} — RM {price:,}" if isinstance(price, (int, float)) else f"    [{i}/{len(car_list)}] ✅ {name}")

        if page < PAGES_TO_SCRAPE:
            print(f"\n  ⏳ Waiting {DELAY_BETWEEN_PAGES}s before next page...")
            time.sleep(DELAY_BETWEEN_PAGES)

    # ── Build Excel ──
    if not all_cars:
        print("\n[!] No cars collected. Exiting.")
        return

    print(f"\n{'='*55}")
    print(f"  📊 Building Excel file...")
    print(f"{'='*55}")

    # Standardize data to the Core Schema
    STANDARD_COLUMNS = [
        "Extraction Time", "Title", "Brand", "Model", "Variant",
        "Year", "Price", "Mileage", "Engine CC", "Transmission",
        "State", "Area", "Listing Link", "Main Image URL"
    ]

    for car in all_cars:
        # Mapping to Standard Core Columns
        if "Car Name" in car: car["Title"] = car.pop("Car Name")
        if "Selling Price (RM)" in car: car["Price"] = car.pop("Selling Price (RM)")
        if "Mileage (km)" in car: car["Mileage"] = car.pop("Mileage (km)")
        if "Engine (L)" in car: car["Engine CC"] = car.pop("Engine (L)")
        if "Detail Page URL" in car: car["Listing Link"] = car.pop("Detail Page URL")
        
        # Ensure all standard columns are present (fill with N/A if missing)
        for col in STANDARD_COLUMNS:
            if col not in car:
                car[col] = "N/A"

    df = pd.DataFrame(all_cars)

    # Reorder: put STANDARD_COLUMNS first, then any remaining columns
    remaining = [c for c in df.columns if c not in STANDARD_COLUMNS]
    df = df[STANDARD_COLUMNS + remaining]
    # Save to Excel with formatting
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Car Specs")

        # Auto-adjust column widths
        worksheet = writer.sheets["Car Specs"]
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(column)),
                df[column].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            adjusted_width = min(max(max_length + 2, 12), 60)
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = adjusted_width

    # ── Final Report ──
    print(f"\n╔══════════════════════════════════════════════════════╗")
    print(f"║   🎉 SCRAPING COMPLETE!                             ║")
    print(f"╠══════════════════════════════════════════════════════╣")
    print(f"║   Total Cars Scraped:     {len(all_cars):<27}║")
    print(f"║   Total Available on Site:{total_available:<27}║")
    print(f"║   Total Columns:          {len(df.columns):<27}║")
    print(f"║   Output File:            {OUTPUT_FILE:<27}║")
    print(f"╚══════════════════════════════════════════════════════╝")

    # Show column summary
    print(f"\n  Columns in Excel ({len(df.columns)} total):")
    for col in df.columns:
        non_na = df[col].apply(lambda x: x != "N/A" and pd.notna(x)).sum()
        print(f"    • {col} ({non_na}/{len(df)} filled)")


if __name__ == "__main__":
    main()
