"""
sebi_fpi_scrape_with_correct_pagination.py

Scrapes SEBI FPI registration cards (0-9, A-Z) including pagination and saves to an Excel file.

Fixes:
 - Uses zero-based page parameter when calling searchFormFpi('n', ...)
 - Parses "1 to 25 of 136 records" to compute per-page and total pages
 - Waits until pagination_inner shows expected range for each page before scraping

Requirements:
  pip install selenium webdriver-manager pandas openpyxl
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import traceback
import math
import re

# ---------- CONFIG ----------
START_URL = "https://www.sebi.gov.in/sebiweb/other/OtherAction.do?doRecognisedFpi=yes&intmId=13"
OUTPUT_XLSX = "sebi_fpi_list_all_pages_corrected.xlsx"
HEADLESS = False          # set True for headless runs
CLICK_DELAY = 0.8         # wait after triggering a change
WAIT_TIMEOUT = 15         # wait for elements to appear
# ----------------------------

COLUMNS = [
    "Name",
    "Registration No.",
    "E-mail",
    "Telephone",
    "Fax No.",
    "Address",
    "Contact Person",
    "Correspondence Address",
    "Validity"
]

TITLE_TO_HEADER = {
    "Name": "Name",
    "Registration No.": "Registration No.",
    "E-mail": "E-mail",
    "Telephone": "Telephone",
    "Fax No.": "Fax No.",
    "Address": "Address",
    "Contact Person": "Contact Person",
    "Correspondence Address": "Correspondence Address",
    "Validity": "Validity",
    # little variants
    "Email": "E-mail",
    "Email ID": "E-mail",
    "Fax": "Fax No.",
    "Registration No": "Registration No.",
}

def make_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1200")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def call_js_safe(driver, script):
    try:
        return driver.execute_script(script)
    except Exception:
        return None

def trigger_letter(driver, letter_id):
    js = f"if (typeof searchFormFpiAlp === 'function') {{ searchFormFpiAlp('{letter_id}'); return true; }} else {{ return false; }}"
    res = call_js_safe(driver, js)
    if res:
        return True
    # fallback
    try:
        el = driver.find_element(By.ID, letter_id)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False

def trigger_page_zero_based(driver, page_zero_based):
    """Call searchFormFpi('n', <zero_based_index>)"""
    js = f"if (typeof searchFormFpi === 'function') {{ searchFormFpi('n', '{page_zero_based}'); return true; }} else {{ return false; }}"
    res = call_js_safe(driver, js)
    if res:
        return True
    # fallback: attempt to click anchor with the matching javascript href (best-effort)
    try:
        anchors = driver.find_elements(By.CSS_SELECTOR, "div.pagination_outer ul li a")
        for a in anchors:
            href = a.get_attribute("href") or ""
            if f"searchFormFpi('n', '{page_zero_based}')" in href or f"searchFormFpi(\"n\", \"{page_zero_based}\")" in href:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
                driver.execute_script("arguments[0].click();", a)
                return True
    except Exception:
        pass
    return False

def parse_pagination_inner(text):
    """
    Parse pagination_inner text like: "1 to 25 of 136 records"
    Returns (start, end, total) as ints or None if cannot parse.
    """
    if not text:
        return None
    # remove non-breaking spaces etc
    txt = text.replace('\xa0', ' ').strip()
    m = re.search(r"(\d+)\s*to\s*(\d+)\s*of\s*(\d+)", txt, re.IGNORECASE)
    if m:
        start = int(m.group(1))
        end = int(m.group(2))
        total = int(m.group(3))
        return (start, end, total)
    return None

def get_total_records_and_perpage(driver):
    """Return (total_records, per_page). If not found, fallback to (1,25)."""
    try:
        p = driver.find_element(By.CSS_SELECTOR, "div.pagination_inner p")
        parsed = parse_pagination_inner(p.text)
        if parsed:
            start, end, total = parsed
            per_page = end - start + 1
            if per_page <= 0:
                per_page = 25
            return total, per_page
    except Exception:
        pass
    # fallback to guessing per_page = 25 and total = 1 page
    return 1, 25

def wait_for_expected_range(driver, expected_start, expected_end, expected_total, timeout=WAIT_TIMEOUT):
    """
    Wait until pagination_inner text contains the expected start (preferred) or expected range.
    Returns True if matched, False if timed out.
    """
    wait = WebDriverWait(driver, timeout)
    def _check(_):
        try:
            p = driver.find_element(By.CSS_SELECTOR, "div.pagination_inner p")
            parsed = parse_pagination_inner(p.text)
            if not parsed:
                return False
            start, end, total = parsed
            # match total and start; sometimes end not exact due to site, so prefer start & total
            return (total == expected_total) and (start == expected_start)
        except Exception:
            return False
    try:
        return wait.until(_check)
    except Exception:
        return False

def scrape_cards_on_current_view(driver):
    rows = []
    card_containers = driver.find_elements(By.CSS_SELECTOR, "div.fixed-table-body.card-table")
    for card in card_containers:
        card_data = {h: "" for h in COLUMNS}
        try:
            card_views = card.find_elements(By.CSS_SELECTOR, "div.card-view")
            for cv in card_views:
                try:
                    title_elem = cv.find_element(By.CSS_SELECTOR, "div.title span")
                    value_elem = cv.find_element(By.CSS_SELECTOR, "div.value span")
                    title = title_elem.text.strip()
                    value = value_elem.text.strip()
                    key = TITLE_TO_HEADER.get(title, None)
                    if key:
                        card_data[key] = value
                    else:
                        t_norm = title.strip().rstrip(':').lower()
                        for k_t, v_t in TITLE_TO_HEADER.items():
                            if k_t.lower().rstrip(':') == t_norm:
                                card_data[v_t] = value
                                break
                except Exception:
                    continue
        except Exception:
            continue
        if card_data.get("Name") or card_data.get("Registration No."):
            rows.append(card_data)
    return rows

def scrape_letter_with_pagination(driver, letter_id):
    collected = []
    ok = trigger_letter(driver, letter_id)
    if not ok:
        print(f"[WARN] couldn't trigger letter {letter_id}")
        return collected

    # Give JS time to load initial results
    time.sleep(CLICK_DELAY)

    # determine total records and per_page from pagination_inner
    total_records, per_page = get_total_records_and_perpage(driver)
    total_pages = max(1, math.ceil(total_records / per_page))
    print(f"[INFO] Letter {letter_id}: detected {total_pages} page(s) (total_records={total_records}, per_page={per_page})")

    for page_num in range(1, total_pages + 1):
        # compute zero-based index for JS call
        zero_idx = page_num - 1

        # If page_num is not 1, call searchFormFpi to move to that page
        if page_num != 1:
            okp = trigger_page_zero_based(driver, zero_idx)
            if not okp:
                print(f"[WARN] could not trigger page {page_num} (zero_idx={zero_idx}) for {letter_id}, attempting to continue")
            # small wait for JS to kick in
            time.sleep(CLICK_DELAY)

        # compute expected start/end for waiting/validation
        expected_start = (page_num - 1) * per_page + 1
        expected_end = min(page_num * per_page, total_records)
        # wait for pagination_inner to show expected start + total (preferred)
        matched = wait_for_expected_range(driver, expected_start, expected_end, total_records, timeout=WAIT_TIMEOUT)
        if not matched:
            # fallback: ensure at least card containers present or continue
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.fixed-table-body.card-table")))
            except Exception:
                pass

        # scrape the cards visible now
        rows = scrape_cards_on_current_view(driver)
        print(f"[INFO] Letter {letter_id} page {page_num}: scraped {len(rows)} rows (expected {expected_end - expected_start + 1})")
        collected.extend(rows)

    # final sanity: if collected < total_records, warn
    if len(collected) != total_records:
        print(f"[WARN] letter {letter_id}: collected {len(collected)} rows but pagination_inner reports {total_records} records")
    return collected

def main():
    letters = ["A1"] + [chr(c) for c in range(ord('A'), ord('Z') + 1)]
    driver = make_driver(headless=HEADLESS)
    try:
        driver.get(START_URL)
        wait = WebDriverWait(driver, WAIT_TIMEOUT)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.clearfix")))
        all_rows = []
        for letter in letters:
            try:
                rows = scrape_letter_with_pagination(driver, letter)
                all_rows.extend(rows)
            except Exception as e:
                print(f"[ERR] Exception scraping letter {letter}: {e}")
                traceback.print_exc()
                continue

        if not all_rows:
            print("[WARN] No rows scraped. Try HEADLESS=False (if not), increase CLICK_DELAY/WAIT_TIMEOUT and re-run.")
            return

        df = pd.DataFrame(all_rows)
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[COLUMNS]
        df = df.drop_duplicates().reset_index(drop=True)
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"[OK] Saved {len(df)} rows to {OUTPUT_XLSX}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
