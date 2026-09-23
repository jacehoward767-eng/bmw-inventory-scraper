"""
BMW Dealership Inventory Scraper (Cloud & Local Edition)
========================================================
Pulls full inventory from BMW of Des Moines and records:
- Vehicle details (year, make, model, trim, color, VIN, price)
- Days on lot (calculated from listing date)
- 7-day view count (scraped from rendered vehicle page via Selenium)
- Photo count and a "Needs Photos?" flag (< 3 images = no real photos)
- A "Changes" sheet comparing today vs the most recent previous report
- Outputs 'inventory.json' for iPhone / Siri / Apple Shortcuts integration

Requirements:
    pip install requests pandas xlsxwriter openpyxl selenium webdriver-manager python-dateutil cloudscraper brotli

Run:
    python maps.py
"""

import os
import re
import time
import glob
import json
import random
from datetime import datetime, timezone
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
BASE_URL         = "https://www.bmwdesmoines.com"
API_URL          = f"{BASE_URL}/api/widget/ws-inv-data/getInventory"
DEALER_SITE_ID   = "bmwofdesmoines"
OUTPUT_FOLDER    = "BMW_Inventory_Reports"
JSON_OUTPUT      = "inventory.json"
PAGE_LOAD_WAIT   = 8      # seconds to wait for JS to render view count
REQUEST_DELAY    = 1.5    # seconds between vehicle page loads (base)
PHOTO_THRESHOLD  = 3      # vehicles with fewer photos than this need shooting

# Residential proxy config — set to None to disable
# Format: "http://user:pass@host:port" or "http://host:port"
PROXY = None

SUV_PREFIXES = ["X1", "X2", "X3", "X4", "X5", "X6", "X7", "XM", "IX"]
# ──────────────────────────────────────────────────────────────────────────────


def jitter_delay(base: float = REQUEST_DELAY):
    """Sleep for base seconds ± up to 40% to avoid rhythmic bot signatures."""
    spread = base * 0.4
    time.sleep(base + random.uniform(-spread, spread))


def build_driver() -> webdriver.Chrome:
    """Builds a headless Chrome instance configured to bypass Akamai bot detection."""
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--remote-debugging-port=9222")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    )
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    if PROXY:
        opts.add_argument(f"--proxy-server={PROXY}")

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=opts)
    
    # Mask navigator.webdriver
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            window.navigator.chrome = { runtime: {} };
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
            Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        """},
    )
    return driver


def fetch_inventory(driver: webdriver.Chrome) -> list:
    """
    Fetches inventory via the DDC API by executing fetch() calls inside
    an active Chrome session to inherit valid Akamai bot-verification cookies.
    """
    print(f"Connecting to {DEALER_SITE_ID} via Chrome session...")
    warm_url = f"{BASE_URL}/new-inventory/"
    print(f"  Warming up browser on {warm_url} ...")
    driver.get(warm_url)
    
    # Allow Akamai scripts to execute and drop tokens
    time.sleep(6)

    full_inventory = []
    seen_vins = set()

    endpoints = [
        {"alias": "INVENTORY_LISTING_DEFAULT_AUTO_NEW",  "config_id": "auto-new"},
        {"alias": "INVENTORY_LISTING_DEFAULT_AUTO_USED", "config_id": "auto-used"},
        {"alias": "SITEBUILDER_RETIRED_SERVICE_LOANERS_1", "config_id": "auto-rsl"},
    ]

    fetch_script = """
    const url = arguments[0];
    const payload = arguments[1];
    const callback = arguments[arguments.length - 1];

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        },
        body: JSON.stringify(payload)
    })
    .then(async response => {
        const text = await response.text();
        try {
            const data = JSON.parse(text);
            callback({status: response.status, data: data});
        } catch(e) {
            callback({status: response.status, rawText: text, error: 'JSON parse error'});
        }
    })
    .catch(err => {
        callback({status: 0, error: err.toString()});
    });
    """

    for ep in endpoints:
        label = ep["config_id"]
        print(f"\n  Fetching {label} inventory...")

        offset = 0
        page_size = 36

        while True:
            print(f"    -> Pulling offset {offset}...", end=" ", flush=True)
            payload = {
                "siteId":               DEALER_SITE_ID,
                "locale":               "en_US",
                "pageAlias":            ep["alias"],
                "widgetName":           "ws-inv-data",
                "inventoryParameters":  {
                    "start": str(offset),
                    "firstRecord": str(offset)
                },
                "preferences": {
                    "pageSize":          str(page_size),
                    "includePricing":    True,
                    "listing.config.id": label,
                },
            }

            success = False
            items = []
            for attempt in range(1, 4):
                try:
                    res = driver.execute_async_script(fetch_script, API_URL, payload)
                    status = res.get("status")
                    if status == 200:
                        items = res.get("data", {}).get("inventory", [])
                        success = True
                        break
                    else:
                        err_msg = res.get("error") or f"HTTP {status}"
                        print(f"(Attempt {attempt} failed: {err_msg})", end=" ")
                        time.sleep(attempt * 4)
                except Exception as e:
                    print(f"(Attempt {attempt} driver error: {e})", end=" ")
                    time.sleep(attempt * 4)

            if not success:
                print("Failed after 3 attempts. Stopping pagination for this category.")
                break

            if items:
                new_vins = [v.get("vin") for v in items if v.get("vin") and v.get("vin") not in seen_vins]

                if not new_vins:
                    print("0 new vehicles (reached end or duplicates).")
                    break

                print(f"Found {len(new_vins)} new vehicles.")

                for v in items:
                    vin = v.get("vin")
                    if vin and vin not in seen_vins:
                        full_inventory.append(v)
                        seen_vins.add(vin)

                if len(items) < page_size:
                    print("    -> End of category reached.")
                    break

                offset += page_size
                jitter_delay(1.5)
            else:
                print("0 vehicles found (End of category).")
                break

    return full_inventory


# ─── SELENIUM / VIEW COUNT ────────────────────────────────────────────────────

def get_view_count(driver: webdriver.Chrome, url: str) -> str:
    try:
        driver.get(url)

        try:
            WebDriverWait(driver, PAGE_LOAD_WAIT).until(
                lambda d: d.execute_script("return document.readyState === 'complete'")
            )
        except Exception:
            pass

        # Strategy 1: Data layer metrics
        try:
            js_views = driver.execute_script(
                "return window.DDC?.trackingData?.recentViews?.total || "
                "window.DDC?.trackingData?.viewCount || "
                "window.digitalData?.page?.pageInfo?.analyticsViews;"
            )
            if js_views and str(js_views).isdigit():
                return str(js_views)
        except Exception:
            pass

        # Strategy 2: Targeted CSS selectors
        for css in [
            ".vdp-analytics-badge",
            "[class*='shopper-activity']",
            "[class*='vehicle-views']",
            "[class*='view-count']",
            "[data-views]"
        ]:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, css)
                for el in elements:
                    text = el.text or el.get_attribute("data-views") or ""
                    num = re.search(r'\d+', text)
                    if num and int(num.group(0)) > 0:
                        return num.group(0)
            except Exception:
                continue

        # Strategy 3: Regex fallbacks across DOM
        source = driver.page_source

        m = re.search(r'"recentViews"\s*:\s*\{\s*"total"\s*:\s*(\d+)', source)
        if m:
            return m.group(1)

        m = re.search(r'"viewCount"\s*:\s*(\d+)', source)
        if m:
            return m.group(1)

        for pattern in [
            r'(\d+)\s+(?:people\s+)?(?:viewed?|views?)\s+(?:this\s+)?(?:vehicle\s+)?in\s+the\s+(?:past|last)\s+7\s+days',
            r'(\d+)\s+views?\s+(?:in\s+)?(?:last|past)\s+7\s+days',
            r'7[- ]day\s+views?[:\s]+(\d+)',
            r'(\d+)\s+7[- ]day\s+views?',
        ]:
            m = re.search(pattern, source, re.IGNORECASE)
            if m:
                return m.group(1)

    except Exception as e:
        print(f"[driver error: {e}]", end=" ")

    return "N/A"


# ─── DATA HELPERS ─────────────────────────────────────────────────────────────

def get_attr(car: dict, name: str, fallback: str = "N/A") -> str:
    for attr in car.get("attributes", []):
        if attr.get("name") == name:
            val = attr.get("value")
            return str(val).strip() if val not in (None, "", "null") else fallback
    return fallback


def get_pricing_field(car: dict) -> str:
    """Extract standard retail price or internet price from DDC pricing blocks."""
    pricing = car.get("pricing", {})
    if isinstance(pricing, dict):
        for key in ["finalPrice", "salePrice", "internetPrice", "retailPrice", "askingPrice"]:
            val = pricing.get(key)
            if val and str(val).replace("$", "").replace(",", "").strip().isdigit():
                return f"${int(float(val)):,}"
    
    attr_price = get_attr(car, "internetPrice", fallback="") or get_attr(car, "retailPrice", fallback="")
    if attr_price and attr_price.replace("$", "").replace(",", "").strip().isdigit():
        return f"${int(float(attr_price)):,}"
    return "Call"


def get_mileage_field(car: dict) -> str:
    """Extract odometer reading."""
    miles = get_attr(car, "odometer", fallback="") or car.get("odometer", "")
    if miles and str(miles).replace(",", "").strip().isdigit():
        return f"{int(float(miles)):,}"
    return "N/A"


def get_photo_info(car: dict) -> tuple[int, str]:
    images = car.get("images", [])
    if isinstance(images, dict):
        images = (
            images.get("images")
            or images.get("imageList")
            or (list(images.values())[0] if images else [])
        )
    photo_count  = len(images) if isinstance(images, list) else 0
    needs_photos = "Yes" if photo_count < PHOTO_THRESHOLD else "No"
    return photo_count, needs_photos


def calculate_days_on_lot(car: dict) -> str:
    raw = car.get("inventoryDate")
    if raw:
        try:
            inv_date = datetime.strptime(str(raw), "%b %d, %Y")
            return str(max(0, (datetime.now() - inv_date).days))
        except ValueError:
            pass
        try:
            from dateutil import parser as dateparser
            listed = dateparser.parse(str(raw))
            if listed:
                if listed.tzinfo is None:
                    listed = listed.replace(tzinfo=timezone.utc)
                return str(max((datetime.now(timezone.utc) - listed).days, 0))
        except Exception:
            pass
    return "N/A"


# ─── PROCESSING ───────────────────────────────────────────────────────────────

def process_data(inventory_list: list, driver: webdriver.Chrome) -> list:
    results = []
    total   = len(inventory_list)

    for i, car in enumerate(inventory_list):
        year  = str(car.get("year",  "")).strip()
        make  = str(car.get("make",  "BMW")).strip()
        model = str(car.get("model", "")).strip()
        trim  = str(car.get("trim",  "")).strip()
        vin   = str(car.get("vin",   "N/A")).strip()

        full_name = " ".join(filter(None, [year, make, model, trim]))
        raw_link  = car.get("link", "")
        full_link = (
            f"{BASE_URL}{raw_link}" if raw_link.startswith("/")
            else (raw_link or f"{BASE_URL}/search/index.htm")
        )

        print(f"[{i+1:>3}/{total}] {full_name} ...", end=" ", flush=True)

        views                     = get_view_count(driver, full_link)
        days_on_lot               = calculate_days_on_lot(car)
        photo_count, needs_photos = get_photo_info(car)
        is_suv                    = "SUV" if any(model.upper().startswith(x) for x in SUV_PREFIXES) else "Car"
        price                     = get_pricing_field(car)
        mileage                   = get_mileage_field(car)

        views_tag  = f"FOUND: {views} views" if views != "N/A" else "NOT FOUND: views"
        days_tag   = f"{days_on_lot}d on lot" if days_on_lot != "N/A" else "lot date unknown"
        photos_tag = f"{photo_count} photos" + (" ⚠" if needs_photos == "Yes" else "")
        print(f"{views_tag}  |  {days_tag}  |  {photos_tag}")

        results.append({
            "Vehicle":       full_name,
            "Year":          year,
            "Make":          make,
            "Model":         model,
            "Trim":          trim,
            "Type":          is_suv,
            "Condition":     str(car.get("condition", "N/A")).title(),
            "Ext Color":     get_attr(car, "exteriorColor"),
            "Int Color":     get_attr(car, "interiorColor"),
            "Price":         price,
            "Mileage":       mileage,
            "Days on Lot":   days_on_lot,
            "Views (7D)":    views,
            "Photos":        photo_count,
            "Needs Photos?": needs_photos,
            "VIN":           vin,
            "Link":          f'=HYPERLINK("{full_link}","View Vehicle")',
        })

        jitter_delay(REQUEST_DELAY)

    return results


def sort_results(results: list) -> list:
    def sort_key(row):
        v = row.get("Views (7D)", "N/A")
        try:
            return (0, -int(v))
        except (ValueError, TypeError):
            return (1, 0)
    return sorted(results, key=sort_key)


# ─── CHANGE DETECTION ─────────────────────────────────────────────────────────

def load_previous_report() -> pd.DataFrame | None:
    pattern = os.path.join(OUTPUT_FOLDER, "BMW_Inventory_*.xlsx")
    files   = sorted(glob.glob(pattern))
    if not files:
        return None
    prev_path = files[-1]
    print(f"  Previous report found: {os.path.basename(prev_path)}")
    try:
        return pd.read_excel(prev_path, sheet_name="Inventory", dtype=str)
    except Exception as e:
        print(f"  Could not load previous report: {e}")
        return None


def build_changes(today: list, prev_df: pd.DataFrame | None) -> dict:
    today_df = pd.DataFrame(today, dtype=str)
    if prev_df is None:
        return {}

    today_vins = set(today_df["VIN"].dropna())
    prev_vins  = set(prev_df["VIN"].dropna())

    new_vins     = today_vins - prev_vins
    new_arrivals = today_df[today_df["VIN"].isin(new_vins)][
        ["Vehicle", "Condition", "Ext Color", "Days on Lot", "Photos", "Needs Photos?", "VIN", "Link"]
    ].copy()

    sold_vins = prev_vins - today_vins
    sold      = prev_df[prev_df["VIN"].isin(sold_vins)][
        ["Vehicle", "Condition", "Ext Color", "Days on Lot", "VIN"]
    ].copy()

    common_vins = today_vins & prev_vins
    merged = today_df[today_df["VIN"].isin(common_vins)][["VIN", "Vehicle", "Needs Photos?", "Photos"]].merge(
        prev_df[prev_df["VIN"].isin(common_vins)][["VIN", "Needs Photos?", "Photos"]],
        on="VIN", suffixes=("_today", "_prev")
    )
    photos_done = merged[
        (merged["Needs Photos?_prev"] == "Yes") & (merged["Needs Photos?_today"] == "No")
    ][["Vehicle", "Photos_prev", "Photos_today", "VIN"]].copy()
    photos_done.columns = ["Vehicle", "Photos Before", "Photos Now", "VIN"]

    views_merged = today_df[today_df["VIN"].isin(common_vins)][["VIN", "Vehicle", "Views (7D)"]].merge(
        prev_df[prev_df["VIN"].isin(common_vins)][["VIN", "Views (7D)"]],
        on="VIN", suffixes=("_today", "_prev")
    )

    def safe_int(val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return None

    views_merged["views_today"] = views_merged["Views (7D)_today"].apply(safe_int)
    views_merged["views_prev"]  = views_merged["Views (7D)_prev"].apply(safe_int)
    views_merged = views_merged.dropna(subset=["views_today", "views_prev"])
    views_merged["Change"] = views_merged["views_today"] - views_merged["views_prev"]

    views_up   = views_merged[views_merged["Change"] > 0].sort_values("Change", ascending=False)
    views_down = views_merged[views_merged["Change"] < 0].sort_values("Change")

    views_up   = views_up[["Vehicle", "views_prev", "views_today", "Change", "VIN"]].copy()
    views_down = views_down[["Vehicle", "views_prev", "views_today", "Change", "VIN"]].copy()
    for df in [views_up, views_down]:
        df.columns = ["Vehicle", "Views Before", "Views Now", "Change", "VIN"]

    return {
        "new_arrivals": new_arrivals,
        "sold":         sold,
        "photos_done":  photos_done,
        "views_up":     views_up,
        "views_down":   views_down,
    }


# ─── EXCEL & JSON EXPORT ──────────────────────────────────────────────────────

def write_changes_sheet(wb, changes: dict, prev_df: pd.DataFrame | None):
    ws = wb.add_worksheet("Changes")

    title_fmt  = wb.add_format({"bold": True, "font_size": 13, "font_color": "#FFFFFF", "bg_color": "#1C6CC6", "valign": "vcenter", "border": 1})
    subhdr_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "valign": "vcenter"})
    cell_fmt   = wb.add_format({"border": 1, "valign": "vcenter"})
    green_fmt  = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#E2EFDA", "font_color": "#375623"})
    red_fmt    = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#FFD7D7", "font_color": "#C00000"})
    gold_fmt   = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#FFEB9C", "font_color": "#9C5700", "bold": True})
    pos_fmt    = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#E2EFDA", "font_color": "#375623", "num_format": '+#,##0;-#,##0'})
    neg_fmt    = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#FFD7D7", "font_color": "#C00000", "num_format": '+#,##0;-#,##0'})

    ws.set_column(0, 0, 38)
    ws.set_column(1, 6, 16)

    row = 0

    if prev_df is None:
        ws.merge_range(row, 0, row, 4, "No previous report found — changes will appear from tomorrow.", title_fmt)
        return

    def write_section(heading: str, df: pd.DataFrame, row_fmt_fn=None):
        nonlocal row
        ws.merge_range(row, 0, row, max(len(df.columns) - 1, 1), heading, title_fmt)
        row += 1
        if df.empty:
            ws.merge_range(row, 0, row, max(len(df.columns) - 1, 1), "None", cell_fmt)
            row += 1
            return
        for c, col in enumerate(df.columns):
            ws.write(row, c, col, subhdr_fmt)
        row += 1
        for _, data_row in df.iterrows():
            for c, (col, val) in enumerate(zip(df.columns, data_row)):
                fmt = row_fmt_fn(col, val) if row_fmt_fn else cell_fmt
                try:
                    ws.write_number(row, c, int(val), fmt)
                except (ValueError, TypeError):
                    ws.write(row, c, str(val) if val is not None else "", fmt)
            row += 1
        row += 1

    def new_fmt(col, val):
        if col == "Needs Photos?":
            return red_fmt if str(val) == "Yes" else green_fmt
        return green_fmt if col == "Vehicle" else cell_fmt

    def views_fmt(col, val):
        if col == "Change":
            try:
                return pos_fmt if int(val) > 0 else neg_fmt
            except (ValueError, TypeError):
                return cell_fmt
        return cell_fmt

    write_section(f"🆕  New Arrivals  ({len(changes['new_arrivals'])} vehicles)",  changes["new_arrivals"],  new_fmt)
    write_section(f"✅  Sold / Removed  ({len(changes['sold'])} vehicles)",         changes["sold"],          lambda col, val: red_fmt if col == "Vehicle" else cell_fmt)
    write_section(f"📸  Photos Completed Since Last Run  ({len(changes['photos_done'])} vehicles)", changes["photos_done"], lambda col, val: gold_fmt if col == "Vehicle" else cell_fmt)
    write_section(f"📈  Views Increased  ({len(changes['views_up'])} vehicles)",    changes["views_up"],      views_fmt)
    write_section(f"📉  Views Decreased  ({len(changes['views_down'])} vehicles)",  changes["views_down"],    views_fmt)


def export_to_excel(results: list, changes: dict, prev_df: pd.DataFrame | None) -> str:
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filepath  = os.path.join(OUTPUT_FOLDER, f"BMW_Inventory_{timestamp}.xlsx")

    df = pd.DataFrame(results)

    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Inventory")

        wb = writer.book
        ws = writer.sheets["Inventory"]

        header_fmt = wb.add_format({"bold": True, "bg_color": "#1C6CC6", "font_color": "#FFFFFF", "border": 1, "align": "center", "valign": "vcenter"})
        cell_fmt   = wb.add_format({"border": 1, "valign": "vcenter"})
        link_fmt   = wb.add_format({"border": 1, "valign": "vcenter", "font_color": "#0563C1", "underline": True})
        num_fmt    = wb.add_format({"border": 1, "valign": "vcenter", "num_format": "#,##0"})
        needs_fmt  = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#FFD7D7", "font_color": "#C00000", "bold": True})
        ok_fmt     = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#E2EFDA", "font_color": "#375623"})
        used_needs_photos_fmt = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": "#FFEB9C", "font_color": "#9C5700", "bold": True})

        for col_num, col_name in enumerate(df.columns):
            ws.write(0, col_num, col_name, header_fmt)

        numeric_cols     = {"Views (7D)", "Days on Lot", "Photos"}
        link_col_idx     = df.columns.get_loc("Link")
        needs_photos_idx = df.columns.get_loc("Needs Photos?")
        condition_idx    = df.columns.get_loc("Condition")

        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            row_condition    = str(row[condition_idx]).lower()
            row_needs_photos = str(row[needs_photos_idx])
            flag_condition   = row_condition == "used" and row_needs_photos == "Yes"

            for col_num, (col_name, value) in enumerate(zip(df.columns, row)):
                if col_num == link_col_idx:
                    url_match = re.search(r'HYPERLINK\("([^"]+)"', str(value))
                    if url_match:
                        ws.write_url(row_num, col_num, url_match.group(1), link_fmt, "View Vehicle")
                    else:
                        ws.write(row_num, col_num, value, cell_fmt)
                elif col_num == needs_photos_idx:
                    ws.write(row_num, col_num, value, needs_fmt if str(value) == "Yes" else ok_fmt)
                elif col_num == condition_idx:
                    ws.write(row_num, col_num, value, used_needs_photos_fmt if flag_condition else cell_fmt)
                elif col_name in numeric_cols:
                    try:
                        ws.write_number(row_num, col_num, int(value), num_fmt)
                    except (ValueError, TypeError):
                        ws.write(row_num, col_num, value, cell_fmt)
                else:
                    ws.write(row_num, col_num, value, cell_fmt)

        num_rows = len(df)
        num_cols = len(df.columns)
        ws.add_table(0, 0, num_rows, num_cols - 1, {
            "name":       "Inventory",
            "style":      "Table Style Medium 2",
            "autofilter": True,
            "columns":    [{"header": c} for c in df.columns],
        })

        col_widths = {
            "Vehicle": 38, "Year": 6, "Make": 8, "Model": 10, "Trim": 18,
            "Type": 6, "Condition": 10, "Ext Color": 18, "Int Color": 18,
            "Price": 12, "Mileage": 12,
            "Days on Lot": 12, "Views (7D)": 12,
            "Photos": 8, "Needs Photos?": 14,
            "VIN": 20, "Link": 16,
        }
        for col_num, col_name in enumerate(df.columns):
            ws.set_column(col_num, col_num, col_widths.get(col_name, 14))

        ws.set_row(0, 20)
        ws.freeze_panes(1, 0)

        write_changes_sheet(wb, changes, prev_df)

    return filepath


def export_to_json(results: list) -> str:
    """Exports a clean JSON file optimized for iPhone Shortcuts / Siri consumption."""
    clean_records = []
    for r in results:
        vin = str(r.get("VIN", "")).strip()
        clean_records.append({
            "yr": str(r.get("Year", "")),
            "make": str(r.get("Make", "")),
            "model": str(r.get("Model", "")),
            "trim": str(r.get("Trim", "")),
            "cond": str(r.get("Condition", "")),
            "color": str(r.get("Ext Color", "")),
            "int": str(r.get("Int Color", "")),
            "price": str(r.get("Price", "Call")),
            "miles": str(r.get("Mileage", "N/A")),
            "views": str(r.get("Views (7D)", "N/A")),
            "photos": r.get("Photos", 0),
            "needs_photos": r.get("Needs Photos?", "No"),
            "dol": str(r.get("Days on Lot", "N/A")),
            "vin6": vin[-6:] if len(vin) >= 6 else vin,
            "vin": vin
        })

    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(clean_records, f, indent=2)

    return JSON_OUTPUT


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    start_time = datetime.now()
    print("=" * 60)
    print("  BMW of Des Moines — Inventory Scraper")
    print(f"  Run time: {start_time.strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # 1. Initialize Headless Browser
    print("\n[1/4] Launching headless browser...")
    driver = build_driver()

    try:
        # 2. Fetch inventory via Chrome-driven API calls
        print("\n[2/4] Fetching inventory from dealer API...")
        inventory = fetch_inventory(driver)

        if not inventory:
            print("\nNo inventory returned. Stopping.")
            return

        print(f"\n  Total vehicles to process: {len(inventory)}")

        # 3. Scrape view counts
        print("\n[3/4] Collecting view counts across vehicle pages...")
        results = process_data(inventory, driver)

    finally:
        driver.quit()

    results = sort_results(results)

    # 4. Compare to previous report
    print("\n[4/4] Comparing to previous report & saving...")
    prev_df = load_previous_report()
    changes = build_changes(results, prev_df)

    if prev_df is not None:
        print(f"  → {len(changes['new_arrivals'])} new arrivals")
        print(f"  → {len(changes['sold'])} sold / removed")
        print(f"  → {len(changes['photos_done'])} vehicles photos completed")
        print(f"  → {len(changes['views_up'])} vehicles views up, "
              f"{len(changes['views_down'])} views down")
    else:
        print("  → No previous report found. Changes sheet will appear from tomorrow.")

    filepath = export_to_excel(results, changes, prev_df)
    json_path = export_to_json(results)

    needs_photos_count = sum(1 for r in results if r["Needs Photos?"] == "Yes")
    print(f"\n  ✓ Excel Report saved: {filepath}")
    print(f"  ✓ Siri JSON saved: {json_path}")
    print(f"  ✓ {len(results)} vehicles recorded.")
    print(f"  ✓ {needs_photos_count} vehicles flagged as needing photos.")

    elapsed = datetime.now() - start_time
    minutes, seconds = divmod(int(elapsed.total_seconds()), 60)
    print(f"  ✓ Total run time: {minutes}m {seconds}s")


if __name__ == "__main__":
    main()
