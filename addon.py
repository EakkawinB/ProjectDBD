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

KEY_DELAY_MS = 60  # ✅ 30 → 60

OUTPUT_DIR = BASE_DIR / "N-AUTHORIZED-CAPITAL"
OUTPUT_FILE = OUTPUT_DIR / "Registered capital.xlsx"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# =====================================================
# POPUP HARD GUARD
# =====================================================

def ensure_popup_closed(page):
    for _ in range(5):
        try:
            if page.locator("#warningModal:visible").count() > 0:
                close_popups(page)
                page.wait_for_timeout(200)  
                continue
            if page.locator(".modal.show").count() > 0:
                close_popups(page)
                page.wait_for_timeout(200)  
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


# =====================================================
# SCRAPER FUNCTIONS
# =====================================================

def get_field(page, label):
    locator = page.locator(
        f"//div[contains(@class,'prompt') and contains(text(),'{label}')]"
        f"/following-sibling::div"
    )
    if locator.count() == 0:
        return None
    try:
        return locator.first.inner_text(timeout=3000).strip()  
    except:
        return None


def get_corporate_status(page):
    try:
        loc = page.locator(
            "//div[contains(@class,'prompt') and contains(text(),'สถานภาพนิติบุคคล')]"
            "/following-sibling::div"
        )
        if loc.count() > 0:
            text = loc.first.inner_text(timeout=3000).strip()  
            if text:
                return text

        for selector in [
            "div.col-6.bold.green",
            "div.col-6.bold.red",
            "div.col-6.bold.yellow",
            "div.col-6.bold.orange",
        ]:
            loc2 = page.locator(selector)
            if loc2.count() > 0:
                text = loc2.first.inner_text(timeout=3000).strip()  
                if text:
                    return text

        loc3 = page.locator("div.col-6.bold")
        for i in range(loc3.count()):
            text = loc3.nth(i).inner_text().strip()
            if 1 <= len(text) <= 50:
                return text

        return None
    except:
        return None


def wait_for_capital(page):
    for _ in range(15):            
        val = get_field(page, "ทุนจดทะเบียน")
        if val:
            return val
        page.wait_for_timeout(300) 
    return None


def scrape_company(page, cname, cid):
    reg = wait_for_capital(page)
    return {
        "company_id": cid,
        "Company Name": cname,
        "Registered Capital": reg if reg else "N/A",
        "Paid-up Capital": get_field(page, "ทุนชำระแล้ว") or "N/A",
        "Corporate Status": get_corporate_status(page) or "N/A"
    }

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
    company_list.append({"company_id": cid, "CompanyName": cname})

TOTAL = len(company_list)
log(f"TOTAL = {TOTAL}")

# =====================================================
# OUTPUT INIT
# =====================================================

done_ids = set()

if OUTPUT_FILE.exists():
    df_existing = pd.read_excel(OUTPUT_FILE)
    df_existing["company_id"] = (
        df_existing["company_id"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.zfill(13)
    )
    done_ids = set(df_existing["company_id"])
    df = df_existing.copy()
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
# SEARCH
# =====================================================

def type_search(page, cid):
    ensure_popup_closed(page)
    popup_gate(page, soft=True)
    close_popups(page)

    input_box = page.locator(
        "input.form-control[placeholder*='ค้นหาด้วยชื่อหรือเลขทะเบียน']"
    )

    input_box.wait_for(state="visible", timeout=5000)   
    input_box.first.scroll_into_view_if_needed()

    for i in range(3):
        try:
            input_box.first.click(timeout=5000, force=True)
            break
        except:
            log(f"[type_search] click retry {i+1}")
            ensure_popup_closed(page)
            popup_gate(page, soft=True)
            close_popups(page)
            page.wait_for_timeout(300)  
    else:
        raise Exception("cannot click input")

    page.evaluate("""
        () => {
            const el = document.querySelector("input.form-control");
            if (el) { el.focus(); el.click(); }
        }
    """)

    page.wait_for_timeout(300)  

    input_box.fill("")
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(200)  

    try:
        input_box.type(cid, delay=KEY_DELAY_MS)
    except:
        input_box.fill(cid)

    page.wait_for_timeout(300)  

    try:
        input_box.press("Enter")
    except:
        page.keyboard.press("Enter")

    page.evaluate("""
        () => {
            const el = document.querySelector("input.form-control");
            if (el) {
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('input', { bubbles: true }));
            }
        }
    """)

    page.wait_for_timeout(500)  

# =====================================================
# FLUSH HELPER
# =====================================================

def flush_buffer(df, buffer):
    if not buffer:
        return df
    df = pd.concat([df, pd.DataFrame(buffer)], ignore_index=True)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    log(f"[FLUSH] บันทึก {len(buffer)} รายการ → {OUTPUT_FILE}")
    buffer.clear()
    return df

# =====================================================
# MAIN
# =====================================================

try:
    with sync_playwright() as p:

        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)

        page.goto(BASE_URL)
        page = safe_wait(page, 1500)  

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

                    log(f"[{idx}] typing {cid}")
                    type_search(page, cid)

                    ensure_popup_closed(page)
                    popup_gate(page, soft=True)
                    close_popups(page)

                    found = False
                    for _ in range(20):             
                        if page.locator("text=ทุนจดทะเบียน").count() > 0:
                            found = True
                            break
                        page.wait_for_timeout(300)  

                    if not found:
                        log(f"[{idx}] NO RESULT")
                        buffer.append({
                            "company_id": cid,
                            "Company Name": rec["CompanyName"],
                            "Registered Capital": "NO RESULT",
                            "Paid-up Capital": "NO RESULT",
                            "Corporate Status": "NO RESULT"
                        })
                        done_ids.add(cid)
                        success = True
                        break

                    page.wait_for_timeout(500)  

                    data = scrape_company(page, rec["CompanyName"], cid)
                    buffer.append(data)
                    done_ids.add(cid)

                    log(f"[{idx}] SUCCESS — {data['Registered Capital']}")

                    if len(buffer) >= 10:
                        df = flush_buffer(df, buffer)

                    success = True
                    break

                except Exception as e:
                    log(f"[{idx}] ERROR retry {attempt+1}: {e}")
                    traceback.print_exc()

                    try:
                        page.goto(BASE_URL)
                        page = safe_wait(page, 1500)   
                    except:
                        page = ensure_page_alive(page, browser)

                    page.wait_for_timeout(500) 
            flush_error_buffer()
            page.wait_for_timeout(200)  

        df = flush_buffer(df, buffer)

        browser.close()
        log("\n===== DONE =====")

except Exception:
    if buffer:
        df = flush_buffer(df, buffer)
    traceback.print_exc()
    flush_error_buffer()
    raise