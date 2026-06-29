"""
Carro.co Deep Car Scraper
=========================
Extracts ALL available car data from Carro's API.

Similar to Carsome, Carro provides most of the detailed fields right in the 
listing API using a pointer-based JSON structure.

Usage: source ../mudah_site/venv/bin/activate && python3 carro_deep_scraper.py
"""

import requests
import pandas as pd
import json
import time
from datetime import datetime

# ──────────────────────────────────────────────
# CONFIGURATION - CHOOSE ONE OPTION
# ──────────────────────────────────────────────

# --- OPTION A: TEST MODE (Quick run for your boss) ---
PAGES_TO_SCRAPE = 2

# --- OPTION B: FULL EXTRACTION MODE (Auto-stops when no cars left) ---
# PAGES_TO_SCRAPE = 99999
CARS_PER_PAGE = 15           # Carro returns 15 cars per page
DELAY_BETWEEN_PAGES = 2      # Seconds between page requests
OUTPUT_FILE = f"Carro-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

URL = 'https://carro.co/_actions/getBuyCarListingData'


# ──────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────

def parse_brand_model(title):
    """
    Carro's API doesn't split brand and model explicitly, so we parse it from the title.
    Example: '2023 HONDA CIVIC RS VTEC 1.5 AT'
    """
    words = title.split()
    brand = 'N/A'
    model = 'N/A'
    
    if words:
        # Remove the year from the front and transmission from the back
        if words[0].isdigit() and len(words[0]) == 4: 
            words.pop(0)
        
        if words and words[-1].upper() in ['AT', 'AUTO', 'MT', 'MANUAL', 'CVT', 'DCT']: 
            words.pop(-1)
            
        two_word_brands = ['LAND ROVER', 'ASTON MARTIN', 'ALFA ROMEO', 'GREAT WALL']
        if len(words) >= 2 and f"{words[0]} {words[1]}".upper() in two_word_brands:
            brand = f"{words[0]} {words[1]}".title()
            words = words[2:]
        elif words:
            brand = words[0].capitalize()
            words.pop(0)
            
        if words: 
            model = " ".join(words)
            
    return brand, model


def extract_car_data(car_item):
    """Extract and organize all fields from a single car's data."""
    title = car_item.get('title', 'N/A')
    
    # Remove year from the beginning of the title if it exists
    words = title.split()
    if words and words[0].isdigit() and len(words[0]) == 4:
        words.pop(0)
    title = " ".join(words) if words else title
    
    brand, model = parse_brand_model(title)
    
    raw_url = car_item.get('detailUrl', '')
    full_link = f"https://carro.co{raw_url}" if raw_url else "N/A"
    
    return {
        # ── IDENTITY ──
        "Extraction Time": datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
        "Car ID": car_item.get('id', 'N/A'),
        "Inventory ID": car_item.get('inventoryId', 'N/A'),
        "Title": title,
        "Brand": brand,
        "Model": model,
        "Year": car_item.get('manufactureYear', 'N/A'),
        
        # ── SPECS ──
        "Engine Capacity": car_item.get('engine_capacity', 'N/A'),
        "Transmission": str(car_item.get('transmissionType', 'N/A')).title(),
        "Fuel Type": str(car_item.get('fuel_type', 'N/A')).capitalize(),
        "Color": str(car_item.get('color', 'N/A')).capitalize(),
        "Mileage (km)": car_item.get('mileage', 'N/A'),
        "Owner Count": car_item.get('owner_count', 'N/A'),
        "Registration Date": car_item.get('original_registration_date', 'N/A'),
        
        # ── PRICING ──
        "Price (Formatted)": car_item.get('price', 'N/A'),
        "Listed Price": car_item.get('listedPrice', 'N/A'),
        "All In Price": car_item.get('allInPrice', 'N/A'),
        "Installment / Month": car_item.get('installmentPerMonth', 'N/A'),
        "Original Price": car_item.get('originalPrice', 'N/A'),
        "Promotional Price": car_item.get('promotionalPrice', 'N/A'),
        
        # ── STATUS & CLASSIFICATION ──
        "Type": str(car_item.get('type', 'N/A')).title(),
        "Certified Status": str(car_item.get('carroCertifiedStatus', 'N/A')).title(),
        "Arrival Status": str(car_item.get('arrivalStatus', 'N/A')).replace('-', ' ').title(),
        "State": "N/A",
        "Area": str(car_item.get('location', 'N/A')).title(),
        
        # ── WARRANTY & IMPORT ──
        "Has Ad Warranty": "Yes" if car_item.get('isAdWarranty') else "No",
        "Valid Agent Warranty": "Yes" if car_item.get('isValidAgentWarranty') else "No",
        "Parallel Import": "Yes" if car_item.get('isParallelImport') else "No",
        
        # ── IMAGES & LINKS ──
        "Main Image URL": car_item.get('imageUrl', 'N/A'),
        "Detail Page URL": full_link,
    }

def scrape_page(page_num):
    """Scrape a single page from Carro's listing API."""
    target_index = (page_num - 1) * CARS_PER_PAGE
    
    payload = {
        "countryCode": "my", "locale": "en", "currentItemIndex": target_index, 
        "filters": {"fuelType": [], "availability": [], "transmission": [], "bodyType": [], "colors": [], "promotions": []},
        "sorting": {"fieldName": "time", "order": "desc"}
    }

    response = requests.post(URL, headers=HEADERS, json=payload, timeout=15)
    
    if response.status_code != 200:
        print(f"  [!] Page {page_num} returned status {response.status_code}")
        return [], 0

    json_data = response.json()
    
    # Resolve the pointer-based JSON format
    root_map = json_data[0]
    car_list_index = root_map.get('carList')
    
    if car_list_index is None:
        return [], 0
        
    car_references = json_data[car_list_index]
    total_cars = root_map.get('totalItems', 0)
    
    cars = []
    for car_pointer in car_references:
        car_format = json_data[car_pointer]
        car_item = {}
        for key, pointer in car_format.items():
            if isinstance(pointer, int) and 0 <= pointer < len(json_data):
                car_item[key] = json_data[pointer]
                
        # Filter out in-feed ads (Ads don't have an ID or Title)
        if 'id' in car_item or 'title' in car_item:
            cars.append(car_item)

    return cars, total_cars

# ──────────────────────────────────────────────
# MAIN SCRAPING LOGIC
# ──────────────────────────────────────────────

def main():
    all_cars = []

    print("╔══════════════════════════════════════════════════════╗")
    print("║   Carro.co Deep Car Scraper                        ║")
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
            name = extracted.get("Title", "Unknown")
            price = extracted.get("Price (Formatted)", "N/A")
            print(f"    [{i}/{len(car_list)}] ✅ {name} — {price}")

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
        if "Price (Formatted)" in car: car["Price"] = car.pop("Price (Formatted)")
        if "Engine Capacity" in car: car["Engine CC"] = car.pop("Engine Capacity")
        if "Mileage (km)" in car: car["Mileage"] = car.pop("Mileage (km)")
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


if __name__ == "__main__":
    main()
