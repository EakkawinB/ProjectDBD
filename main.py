import os
import pandas as pd
from playwright.sync_api import sync_playwright

from config import (
    INPUT_IDS_FILE,
    IDS_SHEET,
    ID_COL,
    BASE_URL,
)

# ===== BROWSER =====
from browser.lifecycle import launch_browser, restart_browser
from browser.readiness import ensure_home_search_ready

# ✅ ใช้ safe function (มี popup auto ในตัว)
from browser.navigation import open_company, safe_action, safe_click

# ===== SCRAPER =====
from scraper.table_scraper import get_years, scrape_table
from scraper.merge import merge_if_updated

# ===== Z-SCORE =====
from zscore.calculator import compute_zscore

# ===== REPORT =====
from reports.zscore_excel import save_zscore_single_sheet
from reports.per_company import save_zscore_per_company

# ===== CHECKPOINT =====
from checkpoint.state import (
    get_done_companies,
    get_fatal_error_companies,
    save_error_company,
    flush_error_buffer,
)

# ===== NOTIFY =====
from notify.outlook import notify_outlook_power_automate

# ===== UTILS =====
from utils.logging_utils import log


def main():

    # =========================
    # ✅ LOAD COMPANY LIST
    # =========================
    df = pd.read_excel(INPUT_IDS_FILE, sheet_name=IDS_SHEET, dtype={ID_COL: str})

    df[ID_COL] = df[ID_COL].astype(str).str.zfill(13)
    df["CompanyName"] = df["CompanyName"].astype(str).str.strip()

    company_list = df[[ID_COL, "CompanyName"]].to_dict("records")

    done = get_done_companies()
    fatal = get_fatal_error_companies()

    log(f"✅ Total companies: {len(company_list)}")

    # =========================
    # ✅ PLAYWRIGHT START
    # =========================
    with sync_playwright() as p:

        browser, page = launch_browser(p)

        for i, rec in enumerate(company_list, start=1):

            cid = rec[ID_COL]
            cname = rec["CompanyName"]

            log(f"❤️ {i}/{len(company_list)} → {cname}")

            # ===== SKIP =====
            if cname in done:
                log("⏭ skip done")
                continue

            if cname in fatal:
                log("⛔ skip fatal")
                continue

            success = False

            # =========================
            # ✅ RETRY LOOP
            # =========================
            for attempt in range(1, 3):

                log(f"🔁 attempt {attempt}/2")

                try:
                    # =====================
                    # HOME READY
                    # =====================
                    ensure_home_search_ready(page, browser)
                    safe_action(page)

                    # =====================
                    # SEARCH COMPANY
                    # =====================
                    open_company(page, cid)
                    safe_action(page)

                    # =====================
                    # BALANCE SHEET
                    # =====================
                    safe_click(page, page.get_by_text("งบแสดงฐานะการเงิน"))
                    page.wait_for_selector("table", timeout=15000)

                    bs_years = get_years(page)
                    bs = scrape_table(page, cid, bs_years)

                    if bs.empty:
                        raise Exception("NO_BALANCE")

                    # =====================
                    # INCOME STATEMENT
                    # =====================
                    safe_click(page, page.get_by_text("งบกำไรขาดทุน"))
                    page.wait_for_selector("table", timeout=15000)

                    is_years = get_years(page)
                    is_df = scrape_table(page, cid, is_years)

                    if is_df.empty:
                        raise Exception("NO_INCOME")

                    # =====================
                    # SAVE RAW DATA
                    # =====================
                    company_dir = os.path.join("output", cid)
                    os.makedirs(company_dir, exist_ok=True)

                    bs, _ = merge_if_updated(
                        os.path.join(company_dir, f"balance_{cid}.xlsx"), bs
                    )

                    is_df, _ = merge_if_updated(
                        os.path.join(company_dir, f"income_{cid}.xlsx"), is_df
                    )

                    # =====================
                    # Z-SCORE
                    # =====================
                    zdf = compute_zscore(bs, is_df, cid, cname)

                    if zdf.empty:
                        raise Exception("Z_EMPTY")

                    # =====================
                    # SAVE RESULT
                    # =====================
                    save_zscore_single_sheet(zdf)
                    save_zscore_per_company(zdf, cid)

                    log(f"✅ SUCCESS: {cname}")

                    success = True
                    break

                except Exception as e:
                    log(f"❌ ERROR: {e}")

                    if attempt == 2:
                        save_error_company(cname, cid, str(e))
                        flush_error_buffer()

                    # ✅ reset page safely
                    try:
                        page.goto(BASE_URL, wait_until="domcontentloaded")
                        page.wait_for_timeout(1000)
                    except:
                        browser, page = restart_browser(p, browser)

            if not success:
                log(f"⚠ FAIL: {cname}")

        # =========================
        # ✅ FINAL STEP
        # =========================
        flush_error_buffer()

        notify_outlook_power_automate(
            subject="[PROD] Z-Score Completed",
            status="COMPLETED"
        )

        browser.close()

        log("🏁 DONE ALL COMPANIES")


# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    main()