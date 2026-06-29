"""
Carlist.my Deep Car Scraper
===========================
Extracts detailed car specifications from Carlist.my using Playwright.

Unlike Carsome and Carro which expose everything in a single API, 
Carlist requires us to:
1. Visit the search listing page to gather car URLs
2. Visit each car's detail page to extract the deep specs 
   (which are neatly organized in HTML span tags).

Usage: source ../mudah_site/venv/bin/activate && python3 carlist_deep_scraper.py
"""

from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
from datetime import datetime

# ──────────────────────────────────────────────
# CONFIGURATION - CHOOSE ONE OPTION
# ──────────────────────────────────────────────

# --- OPTION A: TEST MODE (Quick run for your boss) ---
PAGES_TO_SCRAPE = 1
MAX_CARS_TO_DEEP_SCRAPE = 25

# --- OPTION B: FULL EXTRACTION MODE (Auto-stops when no cars left) ---
# PAGES_TO_SCRAPE = 99999
# MAX_CARS_TO_DEEP_SCRAPE = 999999
OUTPUT_FILE = f"Carlist-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"

USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"

# ──────────────────────────────────────────────
# MAIN SCRAPING LOGIC
# ──────────────────────────────────────────────

def main():
    print("╔══════════════════════════════════════════════════════╗")
    print("║   Carlist.my Deep Car Scraper                      ║")
    print(f"║   Scraping {PAGES_TO_SCRAPE} page(s) (Up to {MAX_CARS_TO_DEEP_SCRAPE} cars)                   ║")
    print("╚══════════════════════════════════════════════════════╝")

    all_cars_data = []
    car_links = []

    with sync_playwright() as p:
        # Launch browser VISIBLY (headless=False) to bypass Cloudflare anti-bot security
        # Cloudflare detects invisible/headless browsers and serves fake 404 pages
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(user_agent=USER_AGENT)
        page = context.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # ── STEP 1: GATHER LINKS ──
        print("\n[STEP 1] Gathering car links from search results...")
        for page_number in range(1, PAGES_TO_SCRAPE + 1):
            print(f"  📋 Skimming Page {page_number}...")
            url = f'https://www.carlist.my/cars-for-sale/malaysia?page_number={page_number}'
            
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                try:
                    page.wait_for_selector('article.listing', timeout=15000)
                except Exception as e:
                    page.screenshot(path=f"carlist_error_page_{page_number}.png")
                    print(f"    [!] Timeout waiting for cars. Saved screenshot to carlist_error_page_{page_number}.png")
                    raise e
                
                soup = BeautifulSoup(page.content(), 'html.parser')
                articles = soup.find_all('article', class_='listing')
                
                # Auto-Stop Safety Logic
                if not articles:
                    print(f"    [!] No cars found on page {page_number}. Auto-stopping pagination.")
                    break
                    
                print(f"    Found {len(articles)} articles on page {page_number}")
                
                for car in articles:
                    link = car.get('data-url', 'N/A')
                    if link != 'N/A':
                        full_link = link if link.startswith('http') else f"https://www.carlist.my{link}"
                        
                        # Grab some surface data while we're here
                        raw_title = car.get('data-title', 'N/A')
                        words = raw_title.split()
                        if words and words[0].isdigit() and len(words[0]) == 4:
                            words.pop(0)
                        clean_title = " ".join(words) if words else raw_title
                        
                        price_tag = car.find('div', class_='listing__price')
                        price = price_tag.text.strip() if price_tag else "N/A"
                        
                        location_icon = car.find('i', class_='icon--location')
                        location = location_icon.parent.text.strip() if location_icon and location_icon.parent else "N/A"
                        
                        location_parts = location.split(' - ')
                        state = location_parts[0].strip() if len(location_parts) > 0 else "N/A"
                        area = location_parts[1].strip() if len(location_parts) > 1 else "N/A"
                        
                        car_links.append({
                            'Extraction Time': datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                            'Listing Link': full_link,
                            'Title': clean_title,
                            'Brand': car.get('data-make', 'N/A'),
                            'Model': car.get('data-model', 'N/A'),
                            'Variant': car.get('data-variant', 'N/A'),
                            'Year': car.get('data-year', 'N/A'),
                            'Price': price,
                            'Installment / Month': car.get('data-installment', 'N/A'),
                            'State': state,
                            'Area': area,
                            'Main Image URL': car.get('data-image-src', 'N/A'),
                        })
                
                time.sleep(random.uniform(2.0, 3.5))
                
            except Exception as e:
                print(f"    [!] Error on page {page_number}: {e}")

        # Limit to the requested max
        car_links = car_links[:MAX_CARS_TO_DEEP_SCRAPE]
        print(f"\n[STEP 2] Deep Scraping {len(car_links)} individual car pages...")
        
        # ── STEP 2: DEEP SCRAPE EACH CAR ──
        for idx, car in enumerate(car_links, 1):
            print(f"  [{idx}/{len(car_links)}] 🚗 {car['Title']}...")
            
            try:
                page.goto(car['Listing Link'], wait_until="domcontentloaded", timeout=45000)
                try:
                    page.wait_for_selector('span.u-width-1\\/2', timeout=10000)
                except Exception:
                    page.wait_for_timeout(3000) # Fallback wait
                
                soup = BeautifulSoup(page.content(), 'html.parser')
                
                # Extract deep specs from pairs of <span class="u-width-1/2">
                spans = soup.find_all('span', class_='u-width-1/2')
                
                deep_specs = {}
                for i in range(0, len(spans) - 1, 2):
                    key = spans[i].get_text(strip=True)
                    val = spans[i+1].get_text(strip=True)
                    
                    # We ignore really long descriptive fields (like full definitions of ABS)
                    if len(val) > 100:
                        val = "Yes"
                        
                    deep_specs[key] = val
                
                # Combine surface data + deep specs
                combined_data = {**car, **deep_specs}
                print(f"    Extracted {len(deep_specs)} deep specs")
                all_cars_data.append(combined_data)
                
                time.sleep(random.uniform(1.0, 2.5))
                
            except Exception as e:
                print(f"    [!] Failed to extract details: {e}")
                all_cars_data.append(car) # Keep the surface data at least

        browser.close()

    # ── STEP 3: EXPORT TO EXCEL ──
    if not all_cars_data:
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

    for car in all_cars_data:
        # Ensure all standard columns are present (fill with N/A if missing)
        # Carlist already extracts most fields with the exact names we need.
        for col in STANDARD_COLUMNS:
            if col not in car:
                car[col] = "N/A"

    df = pd.DataFrame(all_cars_data)
    
    # Reorder: put STANDARD_COLUMNS first, then any remaining columns
    remaining = [c for c in df.columns if c not in STANDARD_COLUMNS]
    df = df[STANDARD_COLUMNS + remaining]

    # Save to Excel with formatting
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Car Specs")

        worksheet = writer.sheets["Car Specs"]
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(column)),
                df[column].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            adjusted_width = min(max(max_length + 2, 12), 50)
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = adjusted_width

    print(f"\n╔══════════════════════════════════════════════════════╗")
    print(f"║   🎉 SCRAPING COMPLETE!                             ║")
    print(f"╠══════════════════════════════════════════════════════╣")
    print(f"║   Total Cars Scraped:     {len(all_cars_data):<27}║")
    print(f"║   Total Columns:          {len(df.columns):<27}║")
    print(f"║   Output File:            {OUTPUT_FILE:<27}║")
    print(f"╚══════════════════════════════════════════════════════╝")

if __name__ == "__main__":
    main()
