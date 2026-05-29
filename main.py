import os
import traceback
from datetime import datetime
from datetime import timedelta
import base64
import time
import pandas as pd     
import requests
from pathlib import Path
from playwright.sync_api import sync_playwright, Error as PlaywrightError
from utils.time_utils import (
    format_elapsed as _format_elapsed_util,
    estimate_avg_time_per_company as _estimate_avg_time_per_company_util,
    estimate_eta as _estimate_eta_util,
)
from utils.text_utils import (
    norm_text as _norm_text_util,
    to_float as _to_float_util,
    looks_like_person_name as _looks_like_person_name_util,
    normalize_company_id as _normalize_company_id_util,
)
from utils.logging_utils import ts as _ts_util, log as _log_util
from utils.zip_utils import zip_and_base64 as _zip_and_base64_util
from utils.excel_utils import safe_to_excel as _safe_to_excel_util
from scraper.merge import merge_if_updated as _merge_if_updated_util
from scraper.table_scraper import get_years as _get_years_util, scrape_table as _scrape_table_util
from scraper.navigation import (
    slow_action as _slow_action_nav,
    slow_click as _slow_click_nav,
    slow_type as _slow_type_nav,
    ensure_page_alive as _ensure_page_alive_nav,
    ensure_home_search_ready as _ensure_home_search_ready_nav,
    home_popup_required as _home_popup_required_nav,
    popup_gate as _popup_gate_nav,
    close_popups as _close_popups_nav,
    handle_optional_popup as _handle_optional_popup_nav,
    wait_table_ready as _wait_table_ready_nav,
    click_tab as _click_tab_nav,
    type_search_with_retry as _type_search_with_retry_nav,
    disable_all_scroll as _disable_all_scroll_nav,
)
from zscore.picker import _pick as _pick_util
from zscore.formatter import _finalize_zdf as _finalize_zdf_util
from zscore.calculator import compute_zscore as _compute_zscore_util
from reports.dashboard import (
    build_zscore_pivot as _build_zscore_pivot_util,
    add_summary_charts as _add_summary_charts_util,
    add_company_zscore_comparison as _add_company_zscore_comparison_util,
    add_dashboard_sheet as _add_dashboard_sheet_util,
    add_company_trend_chart as _add_company_trend_chart_util,
)
from reports.zscore_excel import save_zscore_single_sheet as _save_zscore_single_sheet_util
from reports.per_company import save_zscore_per_company as _save_zscore_per_company_util
from notify.outlook import (
    build_attachments as _build_attachments_util,
    notify_outlook_power_automate as _notify_outlook_power_automate_util,
)
from config import (
    BASE_DIR,
    INPUT_IDS_FILE,
    IDS_SHEET,
    ID_COL,
    BASE_URL,
    DEFAULT_TIMEOUT,
    WATCHDOG_RESTART_EVERY,
    COMPANY_HARD_TIMEOUT,
    OUTPUT_DIR,
    Z_SCORE_FILE,
    Z_SCORE_SHEET,
    HEADLESS,
    MAX_RETRY,
    RUN_START_TIME,
    RUN_END_TIME,
    POWER_AUTOMATE_WEBHOOK,
)

os.chdir(BASE_DIR)

TEST_MODE = False                # True = เทส / False = รันจริง
TEST_LIMIT = None                # จำนวนบริษัทที่จะเทส 
SAFE_MODE = True

def format_elapsed(start_dt, now_dt):
    return _format_elapsed_util(start_dt, now_dt)


def estimate_avg_time_per_company(start_dt, now_dt, processed):
    return _estimate_avg_time_per_company_util(start_dt, now_dt, processed)


def estimate_eta(now_dt, remaining_seconds):
    return _estimate_eta_util(now_dt, remaining_seconds)

COL_ORDER = [
    "บริษัท", "ปี",
    "Current Assets", "Current Liab", "Total Assets", "Retained Earn",
    "Net Income", "Interest Exp", "Tax", "Total Liab",
    "Sales", "Working Cap", "EBIT", "Equity",
    "X1", "X2", "X3", "X4", "X5", "Z'-Score", "สถานะ"
]

MONEY_COLS = [
    "Current Assets", "Current Liab", "Total Assets", "Retained Earn",
    "Net Income", "Interest Exp", "Tax", "Total Liab",
    "Sales", "Working Cap", "EBIT", "Equity"
]
RATIO_COLS = ["X1", "X2", "X3", "X4", "X5"]
Z_COL = "Z'-Score"

def ts():
    return _ts_util()

def log(msg):
    _log_util(msg)


# =========================
# HUMAN LIKE SPEED
# =========================
ACTION_DELAY = 0.15
CLICK_DELAY = 0.25
TYPE_DELAY = 35


def slow_action(delay=ACTION_DELAY):
    return _slow_action_nav(delay=delay)


def slow_click(locator, delay=CLICK_DELAY, force=True):
    return _slow_click_nav(locator, delay=delay, force=force)


def slow_type(locator, text, delay=TYPE_DELAY):
    return _slow_type_nav(locator, text, delay=delay)

# =====================================================
# PAGE SAFETY
# =====================================================

def ensure_page_alive(page, browser):
    return _ensure_page_alive_nav(page, browser)
    
def ensure_home_search_ready(page, browser, timeout=180):
    return _ensure_home_search_ready_nav(page, browser, timeout=timeout)

def norm_text(s):
    return _norm_text_util(s)


def to_float(txt):
    return _to_float_util(txt)
    
def looks_like_person_name(name):
    return _looks_like_person_name_util(name)

def looks_like_university(name):
    if not name:
        return False

    keywords = [
        "มหาวิทยาลัย",
        "University",
        "สถาบัน",
        "Institute",
        "College",
        "ราชภัฏ",
        "อาชีว",
        "วิทยาลัย"
    ]

    name_lower = name.lower()
    return any(k.lower() in name_lower for k in keywords)

def normalize_company_id(cid):
    return _normalize_company_id_util(cid)


def resolve_input_ids_file():
    """Resolve source file with backward-compatible fallback paths."""
    candidates = [
        Path(INPUT_IDS_FILE),
        BASE_DIR / "company_ids.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]

def home_popup_required(page, max_reload=2, timeout=15):
    return _home_popup_required_nav(page, max_reload=max_reload, timeout=timeout)

def popup_gate(page, timeout=10, soft=True):
    return _popup_gate_nav(page, timeout=timeout, soft=soft)


def close_popups(page):
    return _close_popups_nav(page)

def handle_optional_popup(page, timeout=12):
    return _handle_optional_popup_nav(page, timeout=timeout)

def wait_popup_must_appear_then_close(page, appear_timeout=12, close_timeout=15):
    """
    popup ต้องโผล่อย่างน้อย 1 ครั้ง
    ถ้าไม่โผล่ => ห้ามทำขั้นตอนต่อ
    """

    popup_selectors = [
        "[role='dialog']",
        "text=เตือน ผู้ใช้งานระบบ",
        "button:has-text('ยอมรับทั้งหมด')",
        "button:has-text('ยอมรับ')",
        "button:has-text('Accept')",
        "button:has-text('ตกลง')",
        "button:has-text('OK')",
        "button:has-text('ปิด')",
        ".modal",
        ".overlay",
    ]

    start = time.time()
    appeared = False

    # ---------- Phase 1: รอ popup โผล่ ----------
    while time.time() - start < appear_timeout:
        for sel in popup_selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    appeared = True
                    break
            except:
                pass

        if appeared:
            break

        page.wait_for_timeout(200)

    if not appeared:
        raise TimeoutError("❌ Popup did NOT appear (block workflow)")

    log("✔ Popup appeared")

    
    page.wait_for_timeout(300)
    close_popups(page)


    # ---------- Phase 2: ปิด popup ----------
    start = time.time()
    while time.time() - start < close_timeout:
        any_visible = False
        for sel in popup_selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    any_visible = True
                    close_popups(page)
                    page.wait_for_timeout(300)
                    break
            except:
                pass

        if not any_visible:
            log("✔ Popup closed")
            return

    raise TimeoutError("❌ Popup could not be closed")

def wait_critical_popups(page, timeout=6):
    """
    รอเฉพาะ popup ที่บัง interaction (เร็ว)
    """
    start = time.time()

    selectors = [
        "[role='dialog']",
        "button:has-text('ยอมรับทั้งหมด')",
        "button:has-text('ยอมรับ')",
        "button:has-text('Accept')",
        "button:has-text('ตกลง')",
        "button:has-text('OK')",
        "button:has-text('ปิด')",
    ]

    while time.time() - start < timeout:
        found = False
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    found = True
                    close_popups(page)
                    page.wait_for_timeout(150)
                    break
            except:
                pass
        if not found:
            return

    log("⚠ critical popup timeout → continue")

def wait_popups_gone(page, timeout=15):
    """
    รอจน dialog / popup / overlay หาย
    ใช้ time.time() แทน page.timeouts (รองรับ Playwright Python)
    """
    start = time.time()

    possible_overlays = [
        "[role='dialog']",
        "button:has-text('ยอมรับทั้งหมด')",
        "button:has-text('ปิด')",
        "div:has-text('เตือน')",
        "div:has-text('ยอมรับ')",
        "div:has-text('Accept')",
        "div:has-text('ไม่ต้องแสดงผลหน้านี้')",
        ".modal",
        ".overlay",
    ]

    while True:
        any_visible = False

        for sel in possible_overlays:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    any_visible = True
                    close_popups(page)
                    page.wait_for_timeout(300)
                    break
            except:
                pass

        if not any_visible:
            return

        if time.time() - start > timeout:
            log("⚠ popup still visible but timeout reached → continue")
            return

def wait_table_ready(page, timeout=35000):
    close_popups(page)
    page.wait_for_load_state("domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except:
        pass

    selectors = [
        "table thead tr th",
        "table tbody tr",
        "table",
        "[role='table']",
    ]
    last_err = None
    for sel in selectors:
        try:
            page.wait_for_selector(sel, timeout=timeout)
            return
        except Exception as e:
            last_err = e
    if last_err:
        raise last_err
    
    wait_table_ready(page)
    slow_action(0.5)

def click_tab(page, text):
    popup_gate(page, timeout=8, soft=True)
    close_popups(page)
    slow_action(0.3)
    
    tab = page.get_by_text(text, exact=True).first

    tab.wait_for(state="visible", timeout=10000)
    tab.click()
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    page.wait_for_timeout(700)

    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
    except:
        pass

    
    popup_gate(page, timeout=8, soft=True)
    close_popups(page)

def merge_if_updated(old_path, new_df):
    return _merge_if_updated_util(old_path, new_df)


def get_years(page):
    return _get_years_util(page)


def scrape_table(page, cid, years):
    return _scrape_table_util(page, cid, years)

def _pick(df, patterns, label="", prefer_exact=True):
    return _pick_util(df, patterns, label=label, prefer_exact=prefer_exact)

def _finalize_zdf(df):
    return _finalize_zdf_util(df)

def compute_zscore(bs, is_df, company_id, company_name):
    return _compute_zscore_util(bs, is_df, company_id, company_name)


def build_zscore_pivot(z_df):
    return _build_zscore_pivot_util(z_df)




def zip_and_base64(zip_path: Path, files: list[Path]):
    return _zip_and_base64_util(zip_path, files)

def build_latest_zscore_table(z_df):
    """
    ตาราง Z'-Score ล่าสุดต่อบริษัท
    """
    latest = (
        z_df
        .sort_values("ปี")
        .dropna(subset=["Z'-Score"])
        .groupby("บริษัท", as_index=False)
        .last()[["บริษัท", "Z'-Score"]]
        .sort_values("Z'-Score")
    )
    return latest

def save_zscore_single_sheet(new_df):
    return _save_zscore_single_sheet_util(new_df)

def add_summary_charts(workbook):
    return _add_summary_charts_util(workbook)

def add_company_zscore_comparison(workbook):
    return _add_company_zscore_comparison_util(workbook)
    
def add_dashboard_sheet(workbook, pivot_df):
    return _add_dashboard_sheet_util(workbook, pivot_df)


def add_company_trend_chart(ws, start_row, pivot_df):
    return _add_company_trend_chart_util(ws, start_row, pivot_df)

def get_done_companies():
    """
    อ่าน zscore.xlsx เพื่อดูว่าบริษัทไหนทำเสร็จแล้ว
    ใช้คอลัมน์ 'บริษัท' เป็น checkpoint
    """
    if not os.path.exists(Z_SCORE_FILE):
        return set()

    try:
        df = pd.read_excel(Z_SCORE_FILE, sheet_name=Z_SCORE_SHEET)
        if "บริษัท" not in df.columns:
            return set()
        done = set(df["บริษัท"].dropna().astype(str).str.strip().unique())
        log(f"📌 Resume mode: found {len(done)} completed companies")
        return done
    except Exception as e:
        log(f"⚠ Cannot read checkpoint zscore.xlsx → {e}")
        return set()

def summarize_run(total_companies):
    done = len(get_done_companies())
    err = len(get_error_companies())

    log("📌 ===== FINAL SUMMARY =====")
    log(f"Total companies : {total_companies}")
    log(f"Completed       : {done}")
    log(f"Errors          : {err}")
    log("===========================")

    return done, err

ERROR_FILE = os.path.join(OUTPUT_DIR, "error_companies.xlsx")
ERROR_BUFFER = []
RUN_ERROR_SET = set()

def log_data_issue(company_name, company_id, category, note=""):
    ERROR_BUFFER.append({
        "บริษัท": company_name,
        "company_id": company_id,
        "category": category,
        "note": note,
        "time": ts()
    })

def restart_browser(p, browser):
    """
    ♻ Restart browser เพื่อล้าง memory leak
    คืนค่า: browser ใหม่, page ใหม่
    """
    try:
        log("♻ Watchdog: restarting browser to prevent memory leak")
        browser.close()
    except:
        pass

    browser = p.chromium.launch(
        headless=HEADLESS,
        args=[
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage"
        ]
    )
    page = browser.new_page()
    page.set_default_timeout(DEFAULT_TIMEOUT)
    page.goto(BASE_URL, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=15000)
    page.wait_for_timeout(800)

    return browser, page

# =========================
# ERROR CLASSIFICATION
# =========================

# ❌ error ที่ไม่มีทางสำเร็จ → ข้ามถาวร
FATAL_ERRORS = {
    "SEARCH_NOT_FOUND",
    "NO_BALANCE_SHEET",
    "NO_INCOME_STATEMENT",
    "VALID_NO_FINANCIAL",
}

# ✅ error ที่ retry ได้
RETRYABLE_ERRORS = {
    "COMPANY_TIMEOUT",
    "PAGE_TIMEOUT",
    "PLAYWRIGHT_TIMEOUT",
    "TARGET_CLOSED",
    "UNKNOWN_ERROR",
}

def get_error_companies():
    if not os.path.exists(ERROR_FILE):
        return set()

    try:
        df = pd.read_excel(ERROR_FILE)
        if "บริษัท" not in df.columns:
            return set()
        errs = set(df["บริษัท"].dropna().astype(str).str.strip().unique())
        log(f"📌 Resume mode: found {len(errs)} error companies")
        return errs
    except Exception as e:
        log(f"⚠ Cannot read error checkpoint → {e}")
        return set()
    
def get_fatal_error_companies():
    """
    อ่าน error_companies.xlsx
    คืนเฉพาะบริษัทที่เป็น fatal error (ไม่ต้องทำอีก)
    """
    if not os.path.exists(ERROR_FILE):
        return set()

    try:
        df = pd.read_excel(ERROR_FILE)
        if df.empty or "บริษัท" not in df.columns or "error" not in df.columns:
            return set()

        fatal_df = df[df["error"].isin(FATAL_ERRORS)]
        return set(
            fatal_df["บริษัท"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )
    except Exception as e:
        log(f"⚠ Cannot read fatal error list → {e}")
        return set() 

def save_error_company(company_name, company_id, reason):
    ERROR_BUFFER.append({
        "บริษัท": company_name,
        "company_id": company_id,
        "error": reason,
        "time": ts()
    })

    RUN_ERROR_SET.add(company_name)

def flush_error_buffer():
    if not ERROR_BUFFER:
        return

    if os.path.exists(ERROR_FILE):
        df = pd.read_excel(ERROR_FILE)
        df = pd.concat([df, pd.DataFrame(ERROR_BUFFER)], ignore_index=True)
    else:
        df = pd.DataFrame(ERROR_BUFFER)

    df = df.drop_duplicates(["บริษัท"], keep="last")
    safe_to_excel(df, ERROR_FILE)


def read_error_summary(error_file):
    error_file = Path(error_file)
    if not error_file.exists():
        return 0, {}

    df = pd.read_excel(error_file)
    if df.empty:
        return 0, {}

    total_error = len(df)
    by_reason = (
        df["error"].value_counts().to_dict()
        if "error" in df.columns
        else {}
    )

    return total_error, by_reason

def build_mail_body(
    status,
    total,
    completed,
    errors,
    start_time,
    success_rate,
    now_time,
    eta,
    output_path
):
    success_rate = round((completed / total) * 100, 2) if total else 0

    return f"""Status: {status}

Summary
-----------------------
Total companies : {total}
Completed       : {completed}
Errors          : {errors}
Success rate    : {success_rate} %

Timing
-----------------------
Start : {start_time}
Now   : {now_time}

Estimated Finish
-----------------------
ETA   : {eta}

Output
-----------------------
{output_path}
"""

def notify_outlook_batch_report(
    *,
    status,                     # COMPLETED | PAUSED_TIME_WINDOW
    total,
    done,
    errors,
    remaining,
    start_dt,
    end_dt,
    elapsed_time_str,
    avg_time_per_company_sec,
    eta_str,
):
    attachments = build_attachments()

    payload = {
        "subject": (
            "[END OF DAY] Z-Score Batch Paused (Time Window Reached)"
            if status == "PAUSED_TIME_WINDOW"
            else "[PROD] Z-Score Batch Completed"
        ),
        "status": status,

        "progress": {
            "total": total,
            "completed": done,
            "errors": errors,
            "remaining": remaining,
            "progress_pct": round((done / total) * 100, 2) if total else 0,
        },

        "program_timing": {
            "start": start_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "end": end_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "elapsed": elapsed_time_str,
        },

        "performance": {
            "avg_time_per_company_sec": (
                round(avg_time_per_company_sec, 2)
                if avg_time_per_company_sec else None
            )
        },

        "eta": eta_str,
        "attachments": attachments,
    }

    try:
        requests.post(POWER_AUTOMATE_WEBHOOK, json=payload, timeout=30)
        log("📧 Outlook PROD batch report sent (with attachments)")
    except Exception as e:
        log(f"❌ Failed to send Outlook batch report: {e}")


def calc_runtime_info(start_dt, now_dt, completed, total):
    elapsed = now_dt - start_dt

    days = elapsed.days
    hours, rem = divmod(elapsed.seconds, 3600)
    minutes = rem // 60

    elapsed_str = f"{days} days {hours} hours {minutes} minutes"

    # ETA แบบง่าย (ถ้ามี progress)
    if completed > 0:
        avg_sec_per_company = elapsed.total_seconds() / completed
        remaining = max(total - completed, 0)
        eta_dt = now_dt + timedelta(seconds=avg_sec_per_company * remaining)
        eta_str = eta_dt.strftime("%Y-%m-%d %H:%M:%S")
    else:
        eta_str = "-"

    return {
        "now": now_dt.strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed_str": elapsed_str,
        "runtime_days": days,
        "eta": eta_str
    }

def ensure_datetime(t):
    if isinstance(t, str):
        return datetime.fromisoformat(t)
    return t

def notify_outlook_crash(error, start_time):
    attachments = build_attachments()

    payload = {
        "subject": "[CRASH] Z-Score Batch Job Failed",
        "status": "CRASHED",
        "error": str(error),
        "error_type": type(error).__name__,
        "start_time": (
            start_time.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(start_time, "strftime")
            else str(start_time)
        ),
        "crash_time": ts(),
        "attachments": attachments,
    }

    try:
        requests.post(
            POWER_AUTOMATE_WEBHOOK,
            json=payload,
            timeout=20
        )
        log("🚨 CRASH notification sent to Outlook (with attachments)")
    except Exception as e:
        log(f"❌ Failed to send crash notify: {e}")

def file_to_base64(path: Path):
    """
    แปลงไฟล์ -> Base64 สำหรับแนบใน Outlook (Power Automate)
    """
    try:
        if not path.exists():
            return None
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        log(f"⚠ Cannot encode file {path}: {e}")
        return None
    

def build_attachments():
    return _build_attachments_util()

def notify_outlook_power_automate(
    *,
    subject: str,
    status: str,
    payload_extra: dict | None = None
):
    return _notify_outlook_power_automate_util(
        subject=subject,
        status=status,
        payload_extra=payload_extra,
    )

def save_zscore_per_company(zdf, cid):
    return _save_zscore_per_company_util(zdf, cid)

input_ids_path = resolve_input_ids_file()
if not input_ids_path.exists():
    raise FileNotFoundError(
        f"Missing input file: {input_ids_path} (expected 'data/company_ids.xlsx')"
    )

ids_df = pd.read_excel(input_ids_path, sheet_name=IDS_SHEET, dtype={ID_COL: str})

ids_df[ID_COL] = ids_df[ID_COL].astype(str).str.strip().str.zfill(13)
ids_df["CompanyName"] = ids_df["CompanyName"].astype(str).str.strip()

company_list = ids_df[[ID_COL, "CompanyName"]].to_dict("records")

filtered_companies = []

for rec in company_list:
    cname = rec.get("CompanyName", "").strip()

    cid_raw = rec.get(ID_COL)
    if cid_raw is None or pd.isna(cid_raw):
        log_data_issue(cname, cid_raw, "INVALID_NO_ID", "No company_id from source")
        continue

    cid = str(cid_raw).strip()
    if not cid.isdigit() or len(cid) != 13:
        log_data_issue(cname, cid, "INVALID_NO_ID", "Company_id not 13 digits")
        continue

    if looks_like_person_name(cname):
        log_data_issue(cname, cid, "INVALID_PERSON_NAME", "Personal name")
        continue

    if looks_like_university(cname):
        log_data_issue(cname, cid, "INVALID_UNIVERSITY", "University / Institute")
        continue

    filtered_companies.append({
        ID_COL: cid,
        "CompanyName": cname
    })

company_list = filtered_companies
log(f"✅ After pre-filter: {len(company_list)} companies remain")

# =====================================================
# DEDUPLICATE BY COMPANY_ID (ก่อนเปิด Browser)
# =====================================================
dedup_map = {}

for rec in company_list:
    cid = rec[ID_COL]
    cname = rec["CompanyName"]

    if cid not in dedup_map:
        dedup_map[cid] = rec
    else:
        old_name = dedup_map[cid]["CompanyName"]

        if len(cname) > len(old_name):
            dedup_map[cid] = rec

        log_data_issue(
            cname,
            cid,
            "DUPLICATE_ID",
            "Duplicate company_id, merged"
        )

company_list = list(dedup_map.values())
log(f"✅ After ID dedup: {len(company_list)} unique companies remain")

done_companies = get_done_companies()
fatal_error_companies = get_fatal_error_companies()

if TEST_MODE:
    company_list = company_list[:TEST_LIMIT]
    log(f"🧪 TEST MODE → {len(company_list)} companies")

def ensure_company_page(page, cid, timeout=30):
    """
    ทำให้แน่ใจว่าเข้า 'หน้าบริษัท' แล้วจริง
    ถ้ายัง → กด Enter ซ้ำแบบปลอดภัย
    """
    start = time.time()

    while time.time() - start < timeout:
        try:
            # ✅ menu2 มีเฉพาะหน้าบริษัท
            if page.locator("#menu2").count() > 0:
                return True
        except:
            pass

        search = page.locator("form.form-group.search.lg input.form-control").first
        page.evaluate("(el) => el.focus()", search)
        page.keyboard.press("Enter")
        page.wait_for_timeout(800)

    raise TimeoutError(f"Cannot enter company page for {cid}")

def type_search_with_retry(page, browser, cid, retries=2):
    for attempt in range(1, retries + 1):
        log(f"⌨ Search attempt {attempt}/{retries} for {cid}")

        page = ensure_page_alive(page, browser)

        # ✅ รีเซ็ต Home ก่อนทุกครั้ง
        page.goto(BASE_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)

        search = ensure_home_search_ready(page, browser, timeout=180)

        try:
            page.evaluate("(el) => el.focus()", search)
        except:
            pass

        search.fill("")
        page.wait_for_timeout(300)

        slow_type(search, cid)
        page.keyboard.press("Enter")
        slow_action(1.2)

        popup_gate(page, soft=True)

        if page.locator("#menu2").count() > 0:
            log("✅ Entered company page")
            return True

        log("🔁 Not in company page yet")

    raise TimeoutError(f"❌ Cannot search company {cid}")

def clear_error_company(company_name):
    if not os.path.exists(ERROR_FILE):
        return
    df = pd.read_excel(ERROR_FILE)
    df = df[df["บริษัท"] != company_name]
    safe_to_excel(df, ERROR_FILE)

# =========================
# SAFE EXCEL WRITE (WINDOWS)
# =========================
def safe_to_excel(df, path, retries=5, sleep=0.5):
    return _safe_to_excel_util(df, path, retries=retries)

def is_within_run_window():
    now = datetime.now().time()
    return RUN_START_TIME <= now <= RUN_END_TIME


def disable_all_scroll(page):
    page.evaluate("""
        () => {
            // ปิด scroll ทุกชนิด
            window.scrollTo = () => {};
            window.scrollBy = () => {};
            Element.prototype.scrollIntoView = () => {};
            document.documentElement.style.scrollBehavior = "auto";
            document.body.style.scrollBehavior = "auto";
        }
    """)

# =====================================================
# MAIN
# =====================================================

run_start_dt = datetime.now()

try:
    with sync_playwright() as p:
        log("🚀 Launch browser")
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=[
                "--disable-gpu",
                "--no-sandbox",
                "--disable-dev-shm-usage"
            ]
        )
        page = browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)

        page.goto(BASE_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1200)

        disable_all_scroll(page)

        stopped_by_time_window = False

        # =========================
        # MAIN COMPANY LOOP
        # =========================
        for i, rec in enumerate(company_list, start=1):
            # ---- TIME WINDOW CHECK ----
            if not is_within_run_window():
                log("⏰ Outside run window → graceful shutdown")
                stopped_by_time_window = True
                break

            # ---- WATCHDOG ----
            if i % WATCHDOG_RESTART_EVERY == 0:
                browser, page = restart_browser(p, browser)

            log(f"❤️ Progress {i}/{len(company_list)} companies")

            cid = rec[ID_COL]
            cname = rec["CompanyName"]
            company_start_time = time.time()
            # ---- CHECKPOINT ----
            if cname in done_companies:
                log(f"⏭ Skip (already done): {cname}")
                continue

            if cname in fatal_error_companies:
                log(f"⛔ Skip (known error): {cname}")
                continue

            log(f"🔎 Start company_id = {cid} | {cname}")
            success = False

            # =========================
            # PER‑COMPANY RETRY LOOP
            # =========================
            for attempt in range(1, MAX_RETRY + 1):

                if time.time() - company_start_time > COMPANY_HARD_TIMEOUT:
                    log("⏱ Company hard-timeout exceeded → skip")
                    save_error_company(cname, cid, "COMPANY_TIMEOUT")
                    success = True
                    break

                log(f"🔁 Attempt {attempt}/{MAX_RETRY}")

                try:
                    # ---------- STEP 1: HOME / SEARCH ----------
                    home_popup_required(page)
                    popup_gate(page, soft=True)

                    type_search_with_retry(page, browser, cid, retries=2)

                    # ---------- STEP 2: OPEN MENU ----------
                    page.wait_for_selector("#menu2", timeout=25000)
                    slow_click(page.locator("#menu2"))
                    page.wait_for_load_state("domcontentloaded", timeout=20000)
                    page.wait_for_timeout(800) 

                    if SAFE_MODE:
                        handle_optional_popup(page, timeout=8)

                    # ---------- STEP 3: BALANCE SHEET ----------
                    click_tab(page, "งบแสดงฐานะการเงิน")
                    page.wait_for_selector("table", timeout=20000)
                    page.wait_for_timeout(800)

                    bs_years = get_years(page)
                    if not bs_years:
                        save_error_company(cname, cid, "NO_BALANCE_SHEET")
                        flush_error_buffer()
                        success = True
                        break

                    bs = scrape_table(page, cid, bs_years)
                    log(f"📊 BS rows={len(bs)} years={bs_years}")

                    if bs.empty:
                        save_error_company(cname, cid, "VALID_NO_FINANCIAL")
                        flush_error_buffer()
                        success = True
                        break

                    # ---------- STEP 4: INCOME STATEMENT ----------
                    click_tab(page, "งบกำไรขาดทุน")
                    page.wait_for_selector("table", timeout=20000)
                    page.wait_for_timeout(800)

                    is_years = get_years(page)
                    if not is_years:
                        save_error_company(cname, cid, "NO_INCOME_STATEMENT")
                        flush_error_buffer()
                        success = True

                    is_df = scrape_table(page, cid, is_years)
                    log(f"📊 IS rows={len(is_df)} years={is_years}")

                    if is_df.empty:
                        save_error_company(cname, cid, "VALID_NO_FINANCIAL")
                        flush_error_buffer()
                        success = True
                        break

                    # ---------- STEP 5: SAVE RAW ----------
                    cid_dir = OUTPUT_DIR / cid
                    cid_dir.mkdir(exist_ok=True)

                    bs_final, bs_up = merge_if_updated(cid_dir / f"balance_sheet_{cid}.xlsx", bs)
                    if bs_up:
                        safe_to_excel(bs_final, cid_dir / f"balance_sheet_{cid}.xlsx")

                    is_final, is_up = merge_if_updated(cid_dir / f"income_statement_{cid}.xlsx", is_df)
                    if is_up:
                        safe_to_excel(is_final, cid_dir / f"income_statement_{cid}.xlsx")

                    # ---------- STEP 6: Z‑SCORE ----------
                    zdf = compute_zscore(bs_final, is_final, cid, cname)
                    save_zscore_single_sheet(zdf)
                    save_zscore_per_company(zdf, cid)

                    log(f"✅ Done company_id = {cid}")
                    clear_error_company(cname)
                    success = True
                    break

                except Exception as e:
                    err_msg = str(e)
                    err_type = type(e).__name__

                    log(f"❌ Attempt {attempt} Error: {err_type}: {err_msg}")
                    traceback.print_exc()

                    # ===== CASE 1: ค้นหาไม่เจอ → ข้ามถาวร =====
                    if isinstance(e, TimeoutError) and "Cannot search company" in err_msg:
                        save_error_company(cname, cid, "SEARCH_NOT_FOUND")
                        flush_error_buffer()
                        success = True   # ✅ ถือว่าจบบริษัทนี้แล้ว
                        break

                    # ===== CASE 2: หน้าเข้าได้ แต่ไม่มีงบ =====
                    if "NO_BALANCE_SHEET" in err_msg or "NO_INCOME_STATEMENT" in err_msg:
                        save_error_company(cname, cid, "VALID_NO_FINANCIAL")
                        flush_error_buffer()
                        success = True
                        break
                    # ===== CASE 3: timeout / page error → retry ได้ =====
                    if isinstance(e, TimeoutError) or "TargetClosedError" in err_msg:
                        log("🔁 Retryable timeout/page error → retry search")
                        page = ensure_page_alive(page, browser)
                        continue

                    # ===== CASE 4: อื่น ๆ =====
                    if attempt == MAX_RETRY:
                        save_error_company(cname, cid, "MAX_RETRY_EXCEEDED")
                        flush_error_buffer()

                    # ✅ only page‑level recovery here
                    page = ensure_page_alive(page, browser)

            if not success:
                save_error_company(cname, cid, "MAX_RETRY_EXCEEDED")
                flush_error_buffer()

            # ---- RETURN HOME FOR NEXT COMPANY ----
            page.goto(BASE_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(1200)

        # =====================================================
        # 🔴 POST‑BATCH: MUST BE OUTSIDE LOOP
        # =====================================================
        flush_error_buffer()

        done = len(get_done_companies())
        run_end_dt = datetime.now()

        # ===== POST-RUN STATS =====
        errors = len(RUN_ERROR_SET)
        processed = done + errors
        remaining = len(company_list) - processed

        now_dt = run_end_dt

        # ⏱ ใช้ไปแล้วกี่วัน กี่ชั่วโมง
        elapsed_str, _ = format_elapsed(run_start_dt, now_dt)

        # ⚡ เวลาเฉลี่ยต่อบริษัท
        avg_sec = estimate_avg_time_per_company(
            start_dt=run_start_dt,
            now_dt=now_dt,
            processed=processed
        )

        # 📅 ETA
        eta_str = "N/A"
        if avg_sec and remaining > 0:
            eta_dt = estimate_eta(now_dt, avg_sec * remaining)
            eta_str = eta_dt.strftime("%Y-%m-%d %H:%M")

        # ===== FINAL OUTLOOK NOTIFICATION =====
        subject = (
            "[END OF DAY] Z-Score Batch Paused (Time Window Reached)"
            if stopped_by_time_window
            else "[PROD] Z-Score Batch Completed"
        )

        status = (
            "PAUSED_TIME_WINDOW"
            if stopped_by_time_window
            else ("COMPLETED_WITH_ERRORS" if errors > 0 else "COMPLETED")
        )

        notify_outlook_power_automate(
            subject=subject,
            status=status if not TEST_MODE else "TEST",
            
            payload_extra={
                "summary": {
                    "total": len(company_list),
                    "completed": done,
                    "errors": errors,
                    "remaining": remaining,
                    "success_rate": round((done / len(company_list)) * 100, 2)
                    if len(company_list) else 0
                },
                "timing": {
                    "start": run_start_dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "end": run_end_dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "elapsed": elapsed_str,
                },
                "performance": {
                    "avg_time_per_company_sec": round(avg_sec, 2) if avg_sec else None
                },
                "eta": eta_str,
                "mode": "TEST" if TEST_MODE else "PROD",
                "output_path": str(OUTPUT_DIR)
            }
        )

        summarize_run(len(company_list))
        log("🛑 Close browser")
        browser.close()


except Exception as fatal:
    log("🔥 PROGRAM CRASHED (FATAL)")
    traceback.print_exc()

    try:
        flush_error_buffer()
    except:
        pass

    notify_outlook_power_automate(
        subject="[CRASH] Z-Score Batch Job Failed",
        status="CRASHED",
        payload_extra={
            "error": str(fatal),
            "error_type": type(fatal).__name__,
            "start_time": (
                run_start_dt.strftime("%Y-%m-%d %H:%M:%S")
                if hasattr(run_start_dt, "strftime")
                else str(run_start_dt)
            ),
            "crash_time": ts(),
        }
    )

    raise