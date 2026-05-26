# =====================================================
# ENTRYPOINT (CAPITAL SCRAPER - FINAL PRODUCTION FIXED)
# =====================================================

import traceback
import time
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

from config import (
    BASE_URL,
    BASE_DIR,
    HEADLESS,
    DEFAULT_TIMEOUT,
    MAX_RETRY,
    COMPANY_HARD_TIMEOUT,
    RUN_START_TIME,
    RUN_END_TIME,
)

from utils.logging_utils import log
from utils.text_utils import looks_like_person_name
from browser.lifecycle import ensure_page_alive, safe_wait
from browser.popup import home_popup_required, popup_gate, close_popups
from checkpoint.state import flush_error_buffer

KEY_DELAY_MS = 30

OUTPUT_DIR = BASE_DIR / "N-AUTHORIZED-CAPITAL"
OUTPUT_FILE = OUTPUT_DIR / "Registered capital.xlsx"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# =====================================================
# ✅ POPUP HARD GUARD
# =====================================================

def ensure_popup_closed(page):
    for _ in range(5):
        try:
            if page.locator("#warningModal:visible").count() > 0:
                close_popups(page)
                page.wait_for_timeout(300)
                continue

            if page.locator(".modal.show").count() > 0:
                close_popups(page)
                page.wait_for_timeout(300)
                continue

            return
        except:
            return

# =====================================================
# HELPERS
# =====================================================

def is_within_run_window():
    return RUN_START_TIME <= datetime.now().time() <= RUN_END_TIME


def normalize_cid(x):
    return str(x).replace(".0", "").zfill(13)


def wait_for_capital(page):
    for i in range(10):
        val = get_field(page, "ทุนจดทะเบียน")
        if val:
            return val
        page.wait_for_timeout(250)
    return None


# =====================================================
# LOAD DATA
# =====================================================

ids_df = pd.read_excel(BASE_DIR / "data" / "company_ids.xlsx", sheet_name="ids")
ids_df["CompanyName"] = ids_df["CompanyName"].astype(str).str.strip()

company_list = []

for _, row in ids_df.iterrows():
    if pd.isna(row["company_id"]):
        continue

    cid = normalize_cid(row["company_id"])
    cname = str(row["CompanyName"]).strip()

    if looks_like_person_name(cname):
        continue

    company_list.append({
        "company_id": cid,
        "CompanyName": cname
    })

TOTAL = len(company_list)
log(f"TOTAL = {TOTAL}")

# =====================================================
# OUTPUT INIT
# =====================================================

done_ids = set()

if OUTPUT_FILE.exists():
    df = pd.read_excel(OUTPUT_FILE)
    df["company_id"] = df["company_id"].astype(str).str.replace(".0","").str.zfill(13)
    done_ids = set(df["company_id"])
else:
    df = pd.DataFrame(columns=[
        "company_id",
        "Company Name",
        "Registered Capital",
        "Paid-up Capital",
        "Corporate Status"
    ])

buffer = []

# =====================================================
# SEARCH (✅ FIXED)
# =====================================================

def type_search(page, cid):
    input_box = page.locator(
        "input.form-control[placeholder*='ค้นหาด้วยชื่อหรือเลขทะเบียนนิติบุคคล']"
    )

    input_box.wait_for(timeout=8000)

    for i in range(3):
        try:
            ensure_popup_closed(page)
            popup_gate(page, soft=True)
            close_popups(page)

            input_box.first.click(timeout=3000)
            break
        except:
            log(f"retry click ({i+1})")
            ensure_popup_closed(page)
            close_popups(page)
            page.wait_for_timeout(500)
    else:
        raise Exception("click input failed")

    input_box.fill("")
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")

    input_box.type(cid, delay=KEY_DELAY_MS)
    page.wait_for_timeout(150)

    input_box.press("Enter")

# =====================================================
# SCRAPER
# =====================================================

def get_field(page, label):
    locator = page.locator(
        f"//div[contains(@class,'prompt') and contains(text(),'{label}')]/following-sibling::div"
    )

    if locator.count() == 0:
        return None

    try:
        return locator.first.inner_text(timeout=2000).strip()
    except:
        return None


def get_corporate_status(page):
    try:
        loc = page.locator("div.col-6.bold.green, div.col-6.bold.red, div.col-6.bold")
        for i in range(loc.count()):
            text = loc.nth(i).inner_text().strip()
            if len(text) < 50:
                return text
        return None
    except:
        return None


def scrape_company(page, cname, cid):
    reg = wait_for_capital(page)

    if not reg:
        return None

    return {
        "company_id": cid,
        "Company Name": cname,
        "Registered Capital": reg,
        "Paid-up Capital": get_field(page, "ทุนชำระแล้ว"),
        "Corporate Status": get_corporate_status(page)
    }

# =====================================================
# MAIN
# =====================================================

try:
    with sync_playwright() as p:

        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)

        page.goto(BASE_URL)
        page = safe_wait(page, browser, 800)

        for idx, rec in enumerate(company_list, start=1):

            cid = rec["company_id"]
            log(f"\n[{idx}/{TOTAL}] {cid}")

            if cid in done_ids:
                log(f"[{idx}] SKIP")
                continue

            success = False
            start_time = time.time()

            for attempt in range(MAX_RETRY):

                if time.time() - start_time > COMPANY_HARD_TIMEOUT:
                    log(f"[{idx}] TIMEOUT")
                    break

                try:
                    home_popup_required(page)
                    ensure_popup_closed(page)
                    popup_gate(page, soft=True)
                    close_popups(page)

                    type_search(page, cid)

                    page.wait_for_timeout(400)

                    ensure_popup_closed(page)
                    popup_gate(page, soft=True)
                    close_popups(page)

                    found = False
                    for _ in range(12):
                        if page.locator("text=ทุนจดทะเบียน").count() > 0:
                            found = True
                            break
                        page.wait_for_timeout(250)

                    if not found:
                        log(f"[{idx}] NO RESULT")
                        done_ids.add(cid)
                        success = True
                        break

                    data = scrape_company(page, "", cid)

                    if data is None:
                        log(f"[{idx}] NO DATA")
                        done_ids.add(cid)
                        success = True
                        break

                    buffer.append(data)
                    done_ids.add(cid)

                    log(f"[{idx}] SUCCESS")

                    if len(buffer) >= 10:
                        df = pd.concat([df, pd.DataFrame(buffer)], ignore_index=True)
                        df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
                        buffer.clear()

                    success = True
                    break

                except Exception as e:
                    log(f"[{idx}] ERROR retry {attempt+1}")
                    traceback.print_exc()

                    try:
                        page.goto(BASE_URL)
                        page = safe_wait(page, browser, 500)
                    except:
                        page = ensure_page_alive(page, browser)

            flush_error_buffer()
            page.wait_for_timeout(300)

        if buffer:
            df = pd.concat([df, pd.DataFrame(buffer)], ignore_index=True)
            df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

        browser.close()
        log("\n===== DONE =====")

except Exception:
    traceback.print_exc()
    flush_error_buffer()
    raise