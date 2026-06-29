"""
Mudah.my Deep Car Scraper
=========================
Extracts FULL car specifications from mudah.my by:
1. Getting the car listing from the listing page (__NEXT_DATA__)
2. Visiting each car's detail page to extract mcdParams (full specs)
3. Exporting everything to a well-formatted Excel file

Usage: python3 mudah_deep_scraper.py
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import re
import time
import sys
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
DELAY_BETWEEN_REQUESTS = 2 # Seconds between each detail page request (be respectful)
OUTPUT_FILE = f"Mudah-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Setup resilient session with retries
session = requests.Session()
retry = Retry(connect=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)


# ──────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────

def parse_next_data(html_text):
    """Extract the __NEXT_DATA__ JSON from a mudah.my page."""
    soup = BeautifulSoup(html_text, 'html.parser')
    script_tag = soup.find('script', id='__NEXT_DATA__')
    if not script_tag or not script_tag.string:
        return None
    # Handle both formats: raw JSON or window.__NEXT_DATA__ = {...};
    raw = script_tag.string.strip()
    raw = re.sub(r'^window\.__NEXT_DATA__\s*=\s*', '', raw)
    raw = re.sub(r';\s*$', '', raw)
    return json.loads(raw)


def extract_listing_info(car_attrs):
    """Extract basic info from the listing page's car attributes."""
    title = car_attrs.get("subject", "N/A")
    words = title.split()
    if words and words[0].isdigit() and len(words[0]) == 4:
        words.pop(0)
    clean_title = " ".join(words) if words else title

    return {
        "Extraction Time": datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
        "Ad ID": car_attrs.get("adId", car_attrs.get("listId", "N/A")),
        "Title": clean_title,
        "Price": car_attrs.get("priceLabel", car_attrs.get("price", "N/A")),
        "Brand": car_attrs.get("makeName", "N/A"),
        "Model": car_attrs.get("modelName", "N/A"),
        "Year": car_attrs.get("manufacturedYear", "N/A"),
        "Condition": car_attrs.get("conditionName", "N/A"),
        "Transmission (Listing)": car_attrs.get("transmissionName", "N/A"),
        "Car Type": car_attrs.get("carTypeName", "N/A"),
        "State": car_attrs.get("regionName", "N/A"),
        "Area": car_attrs.get("subareaName", "N/A"),
        "Engine Capacity (Listing)": car_attrs.get("engineCapacity", "N/A"),
        "Fuel Type (Listing)": car_attrs.get("fueltype", "N/A"),
        "Seller Name": car_attrs.get("name", "N/A"),
        "Seller Type": "Dealer" if car_attrs.get("companyAd") else "Private",
        "Image Count": car_attrs.get("imageCount", 0),
        "Listing URL": car_attrs.get("adviewUrl", "N/A"),
    }


def extract_deep_specs(detail_attrs):
    """
    Extract ALL specs from the detail page's mcdParams.
    mcdParams is organized by sections: GENERAL, TRANSMISSION, ENGINE,
    DIMENSION & WEIGHT, BRAKES, SUSPENSION, STEERING, TYRES & WHEELS
    """
    specs = {}

    # Extract from mcdParams (the full specs table)
    mcd_params = detail_attrs.get("mcdParams", [])
    for section in mcd_params:
        header = section.get("header", "UNKNOWN")
        params = section.get("params", [])
        for param in params:
            label = param.get("label", "")
            value = param.get("value", "-")
            if label and value != "-":
                specs[label] = value

    # Also extract from categoryParams (basic specs, might overlap)
    cat_params = detail_attrs.get("categoryParams", [])
    for param in cat_params:
        label = param.get("label", "")
        value = param.get("value", "-")
        param_id = param.get("id", "")
        # Only add if not already in specs (mcdParams has more detail)
        if label and value != "-" and label not in specs:
            specs[label] = value

    # Extract seller description / body text
    body = detail_attrs.get("body", "")
    if body:
        # Clean up HTML tags
        body = re.sub(r'<br\s*/?>', '\n', body)
        body = re.sub(r'<[^>]+>', '', body)
        specs["Seller Description"] = body.strip()[:500]  # Limit to 500 chars

    # Extract location label
    specs["Location"] = detail_attrs.get("locationLabel", "N/A")

    # Extract published date
    specs["Published Date"] = detail_attrs.get("publishedDatetime", "N/A")

    return specs


def scrape_listing_page(page_num):
    """Scrape a single listing page and return list of (basic_info, detail_url) tuples."""
    url = f"https://www.mudah.my/malaysia/cars-for-sale?o={page_num}"
    print(f"\n{'='*60}")
    print(f"  Fetching Listing Page {page_num}: {url}")
    print(f"{'='*60}")

    resp = session.get(url, headers=HEADERS, timeout=15)
    if resp.status_code != 200:
        print(f"  [!] Page {page_num} returned status {resp.status_code}")
        return []

    data = parse_next_data(resp.text)
    if not data:
        print(f"  [!] Could not parse __NEXT_DATA__ on page {page_num}")
        return []

    cars_by_id = data.get("props", {}).get("initialState", {}).get("adListing", {}).get("byID", {})
    results = []

    for ad_id, car in cars_by_id.items():
        attrs = car.get("attributes", {})
        basic_info = extract_listing_info(attrs)
        detail_url = attrs.get("adviewUrl", "")
        results.append((basic_info, detail_url))

    print(f"  [+] Found {len(results)} cars on page {page_num}")
    return results


def scrape_detail_page(url):
    """Visit a car's detail page and extract full specs from mcdParams."""
    resp = session.get(url, headers=HEADERS, timeout=15)
    if resp.status_code != 200:
        return {}

    data = parse_next_data(resp.text)
    if not data:
        return {}

    ad_details = data.get("props", {}).get("initialState", {}).get("adDetails", {}).get("byID", {})
    if not ad_details:
        return {}

    # Get the first (and usually only) ad detail
    first_id = list(ad_details.keys())[0]
    detail_attrs = ad_details[first_id].get("attributes", {})
    return extract_deep_specs(detail_attrs)


# ──────────────────────────────────────────────
# MAIN SCRAPING LOGIC
# ──────────────────────────────────────────────

def main():
    all_cars = []
    total_cars = 0
    failed_count = 0

    print("╔══════════════════════════════════════════════════════╗")
    print("║   Mudah.my Deep Car Scraper                        ║")
    print(f"║   Scraping {PAGES_TO_SCRAPE} page(s) with full detail specs        ║")
    print("╚══════════════════════════════════════════════════════╝")

    # Step 1: Collect all cars from listing pages
    all_listings = []
    for page in range(1, PAGES_TO_SCRAPE + 1):
        listings = scrape_listing_page(page)
        
        # Auto-Stop Safety Logic
        if not listings:
            print(f"  [!] No more cars found on page {page}. Auto-stopping pagination.")
            break
            
        all_listings.extend(listings)
        if page < PAGES_TO_SCRAPE:
            time.sleep(DELAY_BETWEEN_REQUESTS)

    total_cars = len(all_listings)
    print(f"\n{'='*60}")
    print(f"  Total cars to process: {total_cars}")
    print(f"  Starting deep extraction (visiting each car page)...")
    print(f"{'='*60}\n")

    # Step 2: Visit each car's detail page for deep specs
    for i, (basic_info, detail_url) in enumerate(all_listings, 1):
        title = basic_info.get("Title", "Unknown")
        progress = f"[{i}/{total_cars}]"

        if not detail_url or detail_url == "N/A":
            print(f"  {progress} ⚠️  Skipping (no URL): {title}")
            basic_info["Deep Specs Status"] = "No URL"
            all_cars.append(basic_info)
            continue

        try:
            print(f"  {progress} 🔍 {title[:60]}...")
            deep_specs = scrape_detail_page(detail_url)

            if deep_specs:
                # Merge basic info with deep specs
                merged = {**basic_info, **deep_specs}
                merged["Deep Specs Status"] = "✅ Success"
                all_cars.append(merged)
                print(f"          ✅ Got {len(deep_specs)} spec fields")
            else:
                basic_info["Deep Specs Status"] = "❌ No specs found"
                all_cars.append(basic_info)
                failed_count += 1
                print(f"          ❌ No deep specs found")

        except Exception as e:
            basic_info["Deep Specs Status"] = f"❌ Error: {str(e)[:50]}"
            all_cars.append(basic_info)
            failed_count += 1
            print(f"          ❌ Error: {e}")

        # Respectful delay between requests
        if i < total_cars:
            time.sleep(DELAY_BETWEEN_REQUESTS)

    # Step 3: Build DataFrame and organize columns
    print(f"\n{'='*60}")
    print(f"  Building Excel file...")
    print(f"{'='*60}")

    # Standardize data to the Core Schema
    STANDARD_COLUMNS = [
        "Extraction Time", "Title", "Brand", "Model", "Variant",
        "Year", "Price", "Mileage", "Engine CC", "Transmission",
        "State", "Area", "Listing Link", "Main Image URL"
    ]

    for car in all_cars:
        # Mapping to Standard Core Columns
        if "Listing URL" in car:
            car["Listing Link"] = car.pop("Listing URL")
        
        # Check if deep specs have the values or fallback to basic specs
        if "Transmission (Listing)" in car and "Transmission" not in car:
            car["Transmission"] = car["Transmission (Listing)"]
        if "Engine Capacity (Listing)" in car and "Engine CC" not in car:
            car["Engine CC"] = car["Engine Capacity (Listing)"]
            
        # Ensure all standard columns are present (fill with N/A if missing)
        for col in STANDARD_COLUMNS:
            if col not in car:
                car[col] = "N/A"

    df = pd.DataFrame(all_cars)

    # Reorder: put STANDARD_COLUMNS first, then any remaining columns
    remaining = [c for c in df.columns if c not in STANDARD_COLUMNS]
    df = df[STANDARD_COLUMNS + remaining]

    # Save to Excel with formatting
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Car Specs')

        # Auto-adjust column widths
        worksheet = writer.sheets['Car Specs']
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(column)),
                df[column].astype(str).str.len().max() if len(df) > 0 else 0
            )
            # Cap at 50 chars width, minimum 12
            adjusted_width = min(max(max_length + 2, 12), 50)
            worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = adjusted_width

    # Final report
    print(f"\n╔══════════════════════════════════════════════════════╗")
    print(f"║   🎉 SCRAPING COMPLETE!                             ║")
    print(f"╠══════════════════════════════════════════════════════╣")
    print(f"║   Total Cars Scraped:     {total_cars:<27}║")
    print(f"║   Successful Deep Specs:  {total_cars - failed_count:<27}║")
    print(f"║   Failed Deep Specs:      {failed_count:<27}║")
    print(f"║   Total Columns:          {len(df.columns):<27}║")
    print(f"║   Output File:            {OUTPUT_FILE:<27}║")
    print(f"╚══════════════════════════════════════════════════════╝")

    # Show a sample of columns found
    print(f"\n  Columns in Excel ({len(df.columns)} total):")
    for col in df.columns:
        non_na = df[col].notna().sum()
        print(f"    • {col} ({non_na}/{len(df)} filled)")


if __name__ == "__main__":
    main()
