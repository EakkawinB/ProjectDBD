import os
import pandas as pd
from playwright.sync_api import sync_playwright

from config import INPUT_IDS_FILE, IDS_SHEET, ID_COL, BASE_URL

from browser.lifecycle import launch_browser, restart_browser
from browser.readiness import ensure_home_search_ready
from browser.navigation import open_company

from browser.popup import popup_gate

from scraper.table_scraper import get_years, scrape_table
from scraper.merge import merge_if_updated

from zscore.calculator import compute_zscore

from reports.zscore_excel import save_zscore_single_sheet
from reports.per_company import save_zscore_per_company

from checkpoint.state import (
    get_done_companies,
    get_fatal_error_companies,
    save_error_company,
    flush_error_buffer,
)

from notify.outlook import notify_outlook_power_automate
from utils.logging_utils import log


# =========================
# ✅ SAFE WRAPPER
# =========================
def safe_click(page, locator):
    popup_gate(page)
    locator.click()
    popup_gate(page)


def safe_step(page):
    popup_gate(page)
    page.wait_for_timeout(300)


# =========================
# ✅ LOAD
# =========================
def load_companies():
    df = pd.read_excel(INPUT_IDS_FILE, sheet_name=IDS_SHEET, dtype={ID_COL: str})
    df[ID_COL] = df[ID_COL].astype(str).str.zfill(13)
    df["CompanyName"] = df["CompanyName"].astype(str).str.strip()
    return df[[ID_COL, "CompanyName"]].to_dict("records")


# =========================
# ✅ PROCESS
# =========================
def process_company(page, browser, p, cid, cname):

    for attempt in range(1, 3):

        try:
            log(f"🔁 attempt {attempt}/2")

            # HOME
            page.goto(BASE_URL, wait_until="domcontentloaded")
            safe_step(page)
            ensure_home_search_ready(page, browser)

            # SEARCH ✅ (แก้หลัก)
            open_company(page, cid)
            safe_step(page)

            # BALANCE
            safe_click(page, page.get_by_text("งบแสดงฐานะการเงิน"))
            page.wait_for_selector("table", timeout=15000)
            safe_step(page)

            bs = scrape_table(page, cid, get_years(page))
            if bs.empty:
                raise Exception("NO_BALANCE")

            # INCOME
            safe_click(page, page.get_by_text("งบกำไรขาดทุน"))
            page.wait_for_selector("table", timeout=15000)
            safe_step(page)

            is_df = scrape_table(page, cid, get_years(page))
            if is_df.empty:
                raise Exception("NO_INCOME")

            # SAVE RAW
            company_dir = os.path.join("output", cid)
            os.makedirs(company_dir, exist_ok=True)

            bs, _ = merge_if_updated(
                os.path.join(company_dir, f"balance_{cid}.xlsx"), bs
            )

            is_df, _ = merge_if_updated(
                os.path.join(company_dir, f"income_{cid}.xlsx"), is_df
            )

            # Z-SCORE
            zdf = compute_zscore(bs, is_df, cid, cname)
            if zdf.empty:
                raise Exception("Z_EMPTY")

            save_zscore_single_sheet(zdf)
            save_zscore_per_company(zdf, cid)

            log(f"✅ SUCCESS: {cname}")
            return True

        except Exception as e:
            log(f"❌ ERROR: {e}")

            if attempt == 2:
                save_error_company(cname, cid, str(e))
                flush_error_buffer()

            try:
                page.goto(BASE_URL)
                popup_gate(page)
            except:
                browser, page = restart_browser(p, browser)

    return False


# =========================
# ✅ MAIN
# =========================
def main():

    company_list = load_companies()
    done = set(get_done_companies())
    fatal = set(get_fatal_error_companies())

    log(f"✅ Total companies: {len(company_list)}")

    with sync_playwright() as p:
        browser, page = launch_browser(p)

        for i, rec in enumerate(company_list, start=1):

            cid = rec[ID_COL]
            cname = rec["CompanyName"]

            log(f"❤️ {i}/{len(company_list)} → {cname}")

            if cname in done:
                log("⏭ skip done")
                continue

            if cname in fatal:
                log("⛔ skip fatal")
                continue

            success = process_company(page, browser, p, cid, cname)

            if not success:
                log(f"⚠ FAIL: {cname}")

        flush_error_buffer()

        notify_outlook_power_automate(
            subject="[PROD] Z-Score Completed",
            status="COMPLETED"
        )

        browser.close()
        log("🏁 DONE ALL COMPANIES")


if __name__ == "__main__":
    main()