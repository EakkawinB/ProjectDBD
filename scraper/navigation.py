import time
from playwright.sync_api import Error as PlaywrightError

from config import BASE_URL, DEFAULT_TIMEOUT
from utils.logging_utils import log


ACTION_DELAY = 0.15
CLICK_DELAY = 0.25
TYPE_DELAY = 35


def slow_action(delay=ACTION_DELAY):
    time.sleep(delay)


def slow_click(locator, delay=CLICK_DELAY, force=True):
    locator.click(force=force)
    time.sleep(delay)


def slow_type(locator, text, delay=TYPE_DELAY):
    locator.type(text, delay=delay)
    time.sleep(ACTION_DELAY)


def ensure_page_alive(page, browser):
    try:
        if page.is_closed():
            raise RuntimeError("page closed")
        page.title()
        return page
    except PlaywrightError as e:
        if "TargetClosedError" in str(e):
            log("🔴 Browser is closed → cannot recover page here")
            raise
        log("♻ Page closed → recreate page")
        page = browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        page.goto(BASE_URL, wait_until="domcontentloaded")
        time.sleep(1.2)
        return page
    except Exception:
        log("♻ Page error → try recreate page")
        page = browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        page.goto(BASE_URL, wait_until="domcontentloaded")
        time.sleep(1.2)
        return page


def popup_gate(page, timeout=10, soft=True):
    start = time.time()
    while time.time() - start < timeout:
        try:
            close_popups(page)
            if page.locator(".modal:visible").count() == 0:
                return
        except Exception:
            pass
        page.wait_for_timeout(500)
    if not soft:
        raise TimeoutError("Popup gate timeout")
    log("⚠ popup gate timeout (soft)")


def close_popups(page):
    popup_selectors = [
        ".modal.show button.btn-close",
        ".modal.show button.close",
        ".swal2-container button.swal2-confirm",
        ".swal2-container button.swal2-cancel",
        ".swal2-container .swal2-close",
    ]
    for sel in popup_selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            try:
                loc.first.click(timeout=800)
                page.wait_for_timeout(200)
            except Exception:
                pass


def home_popup_required(page, max_reload=2, timeout=15):
    for attempt in range(1, max_reload + 1):
        start = time.time()
        while time.time() - start < timeout:
            try:
                popup_gate(page, soft=True)
                close_popups(page)
                time.sleep(0.4)
                if not page.locator(".modal:visible").count():
                    log("✅ Home popup cleared / not blocking")
                    return
            except Exception:
                pass
            time.sleep(0.3)
        log(f"♻ Home popup stuck → reload Home (attempt {attempt})")
        page.goto(BASE_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
    log("⚠ Home popup unresolved → continue (soft)")


def ensure_home_search_ready(page, browser, timeout=180):
    start = time.time()
    last_state = ""
    while time.time() - start < timeout:
        try:
            page = ensure_page_alive(page, browser)
            if BASE_URL not in page.url:
                page.goto(BASE_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)
            popup_gate(page, soft=True)
            close_popups(page)
            page.wait_for_timeout(500)
            for sel in ["input.form-control", "input[type='search']", "form input"]:
                loc = page.locator(sel)
                if loc.count() > 0:
                    box = loc.first
                    if box.is_visible() and box.is_enabled():
                        log(f"✅ Home search ready (selector: {sel})")
                        return box
            last_state = f"url={page.url}"
        except Exception as e:
            last_state = str(e)
        page.wait_for_timeout(2000)
    raise TimeoutError(f"❌ Home search box not available ({last_state})")


def wait_table_ready(page, timeout=35000):
    selectors = ["table thead tr th", "table tbody tr", "table", "[role='table']"]
    last_err = None
    for sel in selectors:
        try:
            page.wait_for_selector(sel, timeout=timeout)
            return
        except Exception as e:
            last_err = e
    if last_err:
        raise last_err


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
    except Exception:
        pass
    popup_gate(page, timeout=8, soft=True)
    close_popups(page)


def handle_optional_popup(page, timeout=12):
    popup_gate(page, timeout=timeout, soft=True)


def type_search_with_retry(page, browser, cid, retries=2):
    for attempt in range(1, retries + 1):
        log(f"⌨ Search attempt {attempt}/{retries} for {cid}")
        page = ensure_page_alive(page, browser)
        page.goto(BASE_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)
        search = ensure_home_search_ready(page, browser, timeout=180)
        try:
            page.evaluate("(el) => el.focus()", search)
        except Exception:
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


def disable_all_scroll(page):
    page.evaluate(
        """
        () => {
            window.scrollTo = () => {};
            window.scrollBy = () => {};
            Element.prototype.scrollIntoView = () => {};
            document.documentElement.style.scrollBehavior = "auto";
            document.body.style.scrollBehavior = "auto";
        }
        """
    )

