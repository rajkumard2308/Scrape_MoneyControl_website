import asyncio
import os
import re
import shutil
import subprocess
import sys
import time
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# SETTINGS
# ============================================================

OUTPUT_FILE = "mutual_funds.xlsx"

# IMPORTANT:
# HEADLESS = True runs Chromium's modern "--headless=new" mode: no
# visible window (works fine on a server/Streamlit Cloud with no
# display, no Xvfb needed), but renders essentially identically to a
# real headed browser. Classic/old headless mode was tried first and
# confirmed to make Moneycontrol's Performance section behave
# differently (no data at all), which is why this isn't just
# headless=True passed straight to Playwright -- see create_browser().
HEADLESS = True

# Moneycontrol + Streamlit Cloud
MAX_CONCURRENCY = 6

PAGE_TIMEOUT = 30000
DATA_TIMEOUT = 10000
ELEMENT_TIMEOUT = 7000

MAX_RETRIES = 2


# ============================================================
# FUNDS
# ============================================================

FUNDS = {
    "Bandhan Small Cap Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-small-cap-fund-regular-plan-growth/MAG2106"
    },

    "Bandhan Large & Mid Cap Fund - Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-large-mid-cap-fund-regular-plan-growth/MAG091"
    },

    "ICICI Prudential Nifty Midcap 150 Index Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/icici-prudential-nifty-midcap-150-index-fund-regular-plan/MPI4540"
    },

    "Bandhan Multi-Factor Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-multi-factor-fund-regular-plan-growth/MAGA093"
    },

    "HDFC Multi-Asset Active FoF Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/hdfc-asset-allocator-fund-of-funds-regular-plan/MHD3469"
    },

    "Kotak Equity Savings Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-equity-savings-fund-regular-plan/MKM887"
    },

    "Kotak Income Plus Arbitrage Omni FOF Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-all-weather-debt-fof-regular-plan/MKM1460"
    },

    "ICICI Prudential Nifty50 Equal Weight Index Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/icici-prudential-nifty50-equal-weight-index-fund-regular-plan-/MPI4613"
    },

    "Kotak Multi Asset Allocation Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-multi-asset-allocation-fund-regular-plan-growth/MKMA059"
    },

    "Nippon India Growth Mid Cap Fund Regular- Growth Plan - Growth Option": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/nippon-india-growth-fund-retail-plan/MRC008"
    },

    "ICICI Prudential Large & Mid Cap Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/icici-prudential-large-mid-cap-fund/MPI003"
    },

    "HDFC Large & Mid Cap Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/hdfc-large-and-mid-cap-fund/MMS001"
    },

    "Nippon India Multi - Asset Omni FoF Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/nippon-india-asset-allocator-fof-regular-plan/MRC2872"
    },

    "Kotak Mid Cap Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-midcap-fund-regular-plan/MKM099"
    },

    "ICICI Prudential Dynamic Asset Allocation Active FOF Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/icici-prudential-asset-allocator-fund/MPI072"
    },

    "Franklin India Opportunities Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/franklin-india-opportunities-fund-regular-plan/MKP028"
    },

    "Franklin India Flexi Cap Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/franklin-india-flexi-cap-fund-regular-plan/MKP003"
    },

    "ICICI Prudential Balanced Advantage Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/icici-prudential-balanced-advantage-fund/MPI126"
    },

    "Kotak Balance Advantage Fund": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-balanced-advantage-fund-regular-plan/MKM1177"
    },

    "Kotak Arbitrage Fund Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/kotak-equity-arbitrage-fund-regular-plan/MKM079"
    },

    "Franklin India Arbitrage Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/franklin-india-arbitrage-fund-regular-plan-growth/MTEA038"
    },

    "HSBC Large and Mid Cap Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/hsbc-large-mid-cap-fund-regular-plan/MHS528"
    },
}


# ============================================================
# Xvfb
# ============================================================

_XVFB_PROCESS = None


def start_virtual_display():
    """
    Start Xvfb ONLY when running headed (HEADLESS = False).

    Headless Chromium does not need a virtual display at all, so with
    the default HEADLESS = True this function is a no-op. It's kept
    around in case someone deliberately flips HEADLESS to False for
    local debugging with a visible-looking browser.
    """

    global _XVFB_PROCESS

    if HEADLESS:
        # No virtual display needed in headless mode.
        return

    # Windows/macOS local machine already has a display.
    if os.name != "posix":
        return

    # If DISPLAY already exists, don't start another one.
    if os.environ.get("DISPLAY"):
        return

    xvfb = shutil.which("Xvfb")

    if not xvfb:
        print(
            "WARNING: Xvfb not found, but HEADLESS=False. "
            "Add 'xvfb' to packages.txt for Streamlit Cloud, "
            "or set HEADLESS = True."
        )
        return

    display = ":99"

    try:

        _XVFB_PROCESS = subprocess.Popen(
            [
                xvfb,
                display,
                "-screen",
                "0",
                "1440x900x24",
                "-ac",
                "-nolisten",
                "tcp",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        os.environ["DISPLAY"] = display

        # Give X server a moment to start.
        import time
        time.sleep(1)

        print(
            f"Xvfb started on DISPLAY={display}"
        )

    except Exception as e:

        print(
            f"WARNING: Could not start Xvfb: {e}"
        )


# Start display as soon as this module loads (no-op when HEADLESS=True).
start_virtual_display()


# ============================================================
# HELPERS
# ============================================================

PERIODS = (
    "1Y",
    "2Y",
    "3Y",
    "5Y",
)


def clean_text(value) -> str:

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
        .replace("\xa0", " ")
    ).strip()


def parse_number(value) -> Optional[float]:

    value = clean_text(value)

    if not value:
        return None

    if value.lower() in {
        "--",
        "-",
        "n/a",
        "na",
        "nil",
        "n.a.",
        "none",
    }:
        return None

    match = re.search(
        r"-?\d[\d,]*(?:\.\d+)?",
        value
    )

    if not match:
        return None

    try:

        return float(
            match.group(0)
            .replace(",", "")
        )

    except Exception:

        return None


def empty_returns() -> Dict[str, Optional[float]]:

    return {
        "1Y": None,
        "2Y": None,
        "3Y": None,
        "5Y": None,
    }


def period_key(value) -> Optional[str]:

    text = clean_text(
        value
    ).lower()

    if re.search(
        r"\b1\s*(?:year|years|yr|yrs|y)\b",
        text
    ):
        return "1Y"

    if re.search(
        r"\b2\s*(?:year|years|yr|yrs|y)\b",
        text
    ):
        return "2Y"

    if re.search(
        r"\b3\s*(?:year|years|yr|yrs|y)\b",
        text
    ):
        return "3Y"

    if re.search(
        r"\b5\s*(?:year|years|yr|yrs|y)\b",
        text
    ):
        return "5Y"

    return None


def has_any_returns(data):

    return any(
        data.get(key) is not None
        for key in PERIODS
    )


def has_all_returns(data):

    return all(
        data.get(key) is not None
        for key in PERIODS
    )


# ============================================================
# PLAYWRIGHT BROWSER INSTALL
# ============================================================

_BROWSER_READY = False


def ensure_playwright_browser():

    global _BROWSER_READY

    if _BROWSER_READY:
        return

    # Playwright's expected Chromium path.
    try:

        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:

            browser_path = p.chromium.executable_path

        if os.path.exists(browser_path):

            print(
                "Playwright Chromium already installed."
            )

            _BROWSER_READY = True
            return

    except Exception:
        pass

    print(
        "Playwright Chromium not found."
    )

    print(
        "Installing Chromium..."
    )

    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "playwright",
            "install",
            "chromium",
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=300,
    )

    print(
        result.stdout
    )

    if result.returncode != 0:

        raise RuntimeError(
            "Could not install Playwright Chromium."
        )

    _BROWSER_READY = True


# ============================================================
# BROWSER
# ============================================================

async def create_browser(playwright):

    args = [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",

        "--disable-gpu",
        "--disable-software-rasterizer",

        "--disable-extensions",

        "--disable-background-networking",
        "--disable-background-timer-throttling",
        "--disable-renderer-backgrounding",

        "--disable-notifications",
        "--disable-popup-blocking",

        "--no-first-run",
        "--no-default-browser-check",

        "--window-size=1440,900",

        "--disable-blink-features=AutomationControlled",
    ]

    # FIX: Chromium's classic --headless mode has a different
    # rendering/fingerprint profile than a real headed browser, and
    # Moneycontrol's Performance section appears to behave differently
    # under it -- this matches exactly what you saw (HEADLESS=True
    # fetched nothing, HEADLESS=False worked).
    #
    # Chromium also has a newer "--headless=new" mode that renders
    # essentially identically to headed Chromium while still running
    # with no visible window (so it still works on a server/Streamlit
    # Cloud with no display, no Xvfb needed). We get that by passing
    # headless=False to Playwright itself (so it doesn't inject its
    # own classic --headless flag) and adding --headless=new to args
    # ourselves instead.
    if HEADLESS:
        args = args + ["--headless=new"]

    browser = await playwright.chromium.launch(
        headless=False,
        args=args,
    )

    return browser


async def create_context(browser):

    context = await browser.new_context(

        viewport={
            "width": 1440,
            "height": 900,
        },

        locale="en-IN",

        timezone_id="Asia/Kolkata",

        user_agent=(
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/131.0.0.0 "
            "Safari/537.36"
        ),

        extra_http_headers={
            "Accept-Language":
                "en-IN,en;q=0.9",

            "Upgrade-Insecure-Requests":
                "1",
        },

        ignore_https_errors=False,
    )

    return context


# ============================================================
# PAGE PREPARATION
# ============================================================

async def prepare_page(page):

    # DO NOT block Moneycontrol resources.
    #
    # The previous implementation blocked resources and
    # could interfere with the site's dynamic frontend.

    async def dismiss_dialog(dialog):

        try:
            await dialog.dismiss()
        except Exception:
            pass

    page.on(
        "dialog",
        lambda dialog: asyncio.create_task(
            dismiss_dialog(dialog)
        )
    )

    # Remove webdriver flag.
    try:

        await page.add_init_script(
            """
            Object.defineProperty(
                navigator,
                'webdriver',
                {
                    get: () => undefined
                }
            );
            """
        )

    except Exception:
        pass


# ============================================================
# PAGE LOAD
# ============================================================

async def load_fund_page(page, url):

    print(
        f"Opening: {url}"
    )

    response = await page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT,
    )

    if response:

        print(
            f"HTTP {response.status}: {url}"
        )

        if response.status >= 400:

            raise RuntimeError(
                f"HTTP {response.status}"
            )
    # Fast load: do not wait for networkidle or fixed multi-second delays.
    try:
        await page.locator("#performance").wait_for(
            state="attached",
            timeout=DATA_TIMEOUT,
        )
    except Exception:
        await page.wait_for_timeout(500)



# ============================================================
# AUM
# ============================================================
#
# Real markup (confirmed on the live site):
#   <li>
#     <span class="OverviewContent_web_name__SnE9y">AUM (Crs.)</span>
#     <span class="OverviewContent_web_value__rmJI_">19,777.42</span>
#   </li>
#
# The trailing hash in each class name (__SnE9y, __rmJI_) is a CSS
# module build hash and can change on redeploy, so we match on the
# stable "OverviewContent_web_name" / "OverviewContent_web_value"
# prefix instead of the exact class.

async def extract_aum(page):

    try:

        name_span = page.locator(
            'span[class*="OverviewContent_web_name"]',
            has_text=re.compile(
                r"^\s*AUM\s*\(Crs?\.?\)\s*$",
                re.I,
            ),
        ).first

        if await name_span.count() > 0:

            value_span = name_span.locator(
                "xpath=following-sibling::span[1]"
            ).first

            if await value_span.count() > 0:

                value = parse_number(
                    await value_span.inner_text(
                        timeout=5000
                    )
                )

                if value is not None:
                    return value

    except Exception as e:
        print(f"WARNING: precise AUM selector failed, falling back: {e}")

    # Fallback: whole-page text regex, kept in case the class names or
    # DOM structure change on Moneycontrol's end.
    try:

        body = clean_text(
            await page.locator(
                "body"
            ).inner_text(
                timeout=5000
            )
        )

        patterns = [

            r"AUM\s*\(Crs?\.?\)\s*"
            r"([\d,]+(?:\.\d+)?)",

            r"AUM\s*\(Cr\.?\)\s*"
            r"([\d,]+(?:\.\d+)?)",

            r"\bAUM\s+"
            r"([\d,]+(?:\.\d+)?)",
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                body,
                re.I
            )

            if match:

                value = parse_number(
                    match.group(1)
                )

                if value is not None:
                    return value

    except Exception:
        pass

    return None


# ============================================================
# PERFORMANCE SECTION
# ============================================================

async def get_performance_section(page):

    performance = page.locator(
        "#performance"
    ).first

    try:

        await performance.wait_for(
            state="attached",
            timeout=DATA_TIMEOUT
        )

        await performance.scroll_into_view_if_needed(
            timeout=5000
        )

        return performance

    except Exception:
        pass

    return None


# ============================================================
# SIP CLICK
# ============================================================


async def click_sip(page):
    """
    Click the real SIP toggle and VERIFY that Moneycontrol switched the
    performance table to SIP mode before returning.

    Confirmed live markup:
        //*[@id="performance"]/div/div[2]/div/label[2]
    is the "SIP" label (label[1] = Lumpsum, label[2] = SIP, label[3] =
    Quarterly, label[4] = Best & Worst). Each label wraps a styled
    fake-radio control:
        <span class="Performance_web_customRadio__BpFF7"></span>

    IMPORTANT: this page also has ad content (banners/iframes) that can
    sit on top of other elements. Playwright's simulated mouse click
    (even with force=True) still fires at the element's real screen
    coordinates, so an overlapping ad can silently swallow the click.
    To rule that out entirely, this now calls .click() on the element
    directly inside the page's own JS context (locator.evaluate(el =>
    el.click())) -- a genuine, bubbling click event with no on-screen
    coordinates involved, so nothing can visually intercept it. Real
    mouse clicks are kept only as a last-resort fallback.
    """

    performance = page.locator("#performance").first

    clicked = False

    # Primary: your confirmed XPath for the SIP label itself.
    try:
        sip_control = page.locator(
            'xpath=//*[@id="performance"]/div/div[2]/div/label[2]'
        ).first

        if await sip_control.count() > 0:
            await sip_control.evaluate("el => el.click()")
            clicked = True

    except Exception as e:
        print(f"WARNING: SIP xpath JS-click failed: {e}")

    if not clicked:
        # Fallback: JS-native click directly on the custom-radio span
        # inside that label. The trailing hash in the class name
        # (__BpFF7) is a CSS module build hash and can change on
        # redeploy, so match by prefix.
        try:
            sip_span = page.locator(
                '#performance span[class*="Performance_web_customRadio"]'
            ).nth(1)  # 0-indexed: Lumpsum=0, SIP=1, Quarterly=2, Best&Worst=3

            if await sip_span.count() > 0:
                await sip_span.evaluate("el => el.click()")
                clicked = True

        except Exception as e:
            print(f"WARNING: SIP span JS-click failed: {e}")

    if not clicked:
        # Last resort: a real simulated mouse click via Playwright.
        try:
            sip_label = performance.get_by_text("SIP", exact=True).first

            if await sip_label.count() > 0:
                await sip_label.scroll_into_view_if_needed(timeout=3000)
                await sip_label.click(timeout=5000, force=True)
                clicked = True

        except Exception as e:
            print(f"WARNING: SIP text fallback click failed: {e}")

    if not clicked:
        print("WARNING: Could not click the SIP control at all")
        return False

    # Verify the table ACTUALLY switched by polling its real headers --
    # the exact same check extract_sip() itself uses (table_mode_matches)
    # -- rather than the <h2> title text. Logs confirmed a real race:
    # the title updates to "SIP Returns" before the <table> body has
    # re-rendered with SIP's columns (Invested/Latest/Start Date), so
    # the old title-only check reported success too early and
    # extract_sip() then read a table that didn't match SIP's shape.
    deadline = time.monotonic() + (DATA_TIMEOUT / 1000)

    while time.monotonic() < deadline:

        try:
            tables = performance.locator("table")
            count = await tables.count()

            for i in range(count):
                rows = await read_table(tables.nth(i))

                if table_mode_matches(rows, "SIP"):
                    return True

        except Exception:
            pass

        await page.wait_for_timeout(250)

    try:
        snippet = await performance.inner_text(timeout=2000)
        snippet = re.sub(r"\s+", " ", snippet)[:300]
    except Exception:
        snippet = "<unavailable>"

    # Dump the actual HTML of the toggle group so a persistent
    # failure can be diagnosed from real evidence instead of
    # another guess -- this will show any real <input> element,
    # onClick wiring, or structural difference we haven't seen yet.
    try:
        toggle_html = await page.locator(
            'xpath=//*[@id="performance"]/div/div[2]/div'
        ).first.evaluate("el => el.outerHTML")
        toggle_html = toggle_html[:1500]
    except Exception:
        toggle_html = "<unavailable>"

    print(
        f"WARNING: SIP table never matched the expected header "
        f"structure within {DATA_TIMEOUT}ms | "
        f"performance section snippet: {snippet!r} | "
        f"toggle group HTML: {toggle_html!r}"
    )
    return False

# ============================================================
# TABLE READER
# ============================================================

async def read_table(table):
    try:
        return await table.evaluate(
            """
            table => Array.from(table.querySelectorAll('tr')).map(
                tr => Array.from(tr.querySelectorAll('th, td'))
                    .map(td => (td.innerText || '').replace(/\\s+/g, ' ').trim())
            ).filter(row => row.length)
            """
        )
    except Exception:
        return []


# ============================================================
# TABLE PARSER
# ============================================================


def normalize_header(value) -> str:
    """Normalize table headers for reliable matching."""
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        clean_text(value).lower()
    ).strip()


def find_annualised_column(rows) -> Optional[int]:
    """
    Find the EXACT Annualised column from the table header.

    We never infer the column from its position when the header exists.
    This is important because Lumpsum and SIP have different columns.
    """
    for row in rows[:3]:
        for index, value in enumerate(row):
            header = normalize_header(value)

            if header in {
                "annualised",
                "annualized",
            } or "annualised" in header or "annualized" in header:
                return index

    return None


def table_mode_matches(rows, mode: str) -> bool:
    """
    Confirm that the table is the requested Moneycontrol performance mode.
    """
    if not rows:
        return False

    headers = {
        normalize_header(value)
        for row in rows[:3]
        for value in row
    }

    has_annualised = any(
        "annualised" in h or "annualized" in h
        for h in headers
    )

    if not has_annualised:
        return False

    if mode == "Lumpsum":
        # Moneycontrol Lumpsum table:
        # Period | Absolute | Annualised | Category Average | Rank
        has_absolute = any("absolute" in h for h in headers)
        has_category = any("category average" in h for h in headers)
        return has_absolute and has_category

    if mode == "SIP":
        # Moneycontrol SIP table:
        # Period | Rs.1000 SIP Start Date | Invested | Latest |
        # Absolute | Annualised
        has_invested = any("invested" in h for h in headers)
        has_latest = any("latest" in h for h in headers)
        has_start_date = any(
            "sip start date" in h or "start date" in h
            for h in headers
        )
        return has_invested and has_latest and has_start_date

    return False


async def parse_table(table, mode):
    """
    Extract ONLY Annualised(%) for 1Y, 2Y, 3Y and 5Y.

    Critical rule:
      - The Annualised column is located from the actual <th>.
      - A '--' / None remains None.
      - We NEVER substitute Absolute, Category Average, Rank,
        Invested, Latest, or values from another table.
    """
    rows = await read_table(table)
    result = empty_returns()

    if not rows:
        return result

    # Reject unrelated tables. This is the key protection against
    # accidentally reading the Yearly Returns table or stale Lumpsum data.
    if not table_mode_matches(rows, mode):
        return result

    annualised_index = find_annualised_column(rows)

    if annualised_index is None:
        return result

    for row in rows:
        if not row:
            continue

        key = period_key(row[0])

        if key not in PERIODS:
            continue

        if annualised_index >= len(row):
            continue

        # If Moneycontrol displays "--", parse_number returns None.
        # Keep it as None; do NOT search another column.
        result[key] = parse_number(row[annualised_index])

    return result

# ============================================================
# LUMPSUM
# ============================================================


async def extract_lumpsum(page, performance):
    """
    Extract Lumpsum Annualised(%) values for 1Y/2Y/3Y/5Y only.

    This is the default view of #performance -- no click needed. It
    just waits for the "Absolute and Annualised Returns" table to be
    ready, then reads it.
    """
    result = empty_returns()

    try:
        await page.wait_for_function(
            r"""
            () => {
                const el = document.querySelector("#performance");
                if (!el) return false;

                const text = (el.innerText || "")
                    .replace(/\s+/g, " ")
                    .toLowerCase();

                return (
                    text.includes("absolute and annualised returns") ||
                    text.includes("absolute and annualized returns")
                );
            }
            """,
            timeout=DATA_TIMEOUT,
        )
    except Exception:
        pass

    tables = performance.locator("table")
    count = await tables.count()

    for i in range(count):
        try:
            table = tables.nth(i)
            rows = await read_table(table)

            if not table_mode_matches(rows, "Lumpsum"):
                continue

            parsed = await parse_table(table, "Lumpsum")

            for key in PERIODS:
                # Only copy the exact Annualised value. None stays None.
                result[key] = parsed[key]

            # Return once the correct table has been found.
            # We intentionally do NOT use another table as a fallback.
            return result

        except Exception as e:
            print(f"WARNING: Lumpsum table {i} parse failed: {e}")
            continue

    print("WARNING: Correct Lumpsum performance table not found")
    return result



async def extract_sip(page, performance):
    """
    Extract SIP Annualised(%) values for 1Y/2Y/3Y/5Y only.

    The SIP table has a different layout from Lumpsum, so this function
    first switches to SIP (see click_sip()) and verifies the real "SIP
    Returns" table title before reading -- no "Performance" tab click
    is needed beforehand; #performance is already present on the page
    with Lumpsum showing by default.
    """
    result = empty_returns()

    clicked = await click_sip(page)

    if not clicked:
        print("WARNING: SIP mode could not be verified")
        return result

    tables = performance.locator("table")
    count = await tables.count()

    for i in range(count):
        try:
            table = tables.nth(i)
            rows = await read_table(table)

            if not table_mode_matches(rows, "SIP"):
                continue

            parsed = await parse_table(table, "SIP")

            for key in PERIODS:
                # Only copy the exact Annualised value. None stays None.
                result[key] = parsed[key]

            # Correct SIP table found. Never fall back to Lumpsum/text.
            return result

        except Exception as e:
            print(f"WARNING: SIP table {i} parse failed: {e}")
            continue

    print("WARNING: Correct SIP performance table not found")
    return result


# ============================================================
# SINGLE FUND
# ============================================================

async def scrape_one(
    context,
    semaphore,
    name,
    mode
):
    """
    mode: "Lumpsum", "SIP", or "Both".

    "Both" loads the fund's page ONCE and extracts Lumpsum then SIP
    from that same page, instead of loading the page twice (once per
    mode). This is the main speed win when a user wants both sets of
    numbers for several funds at once.
    """

    async with semaphore:

        url = FUNDS[name]["url"]
        want_lumpsum = mode in ("Lumpsum", "Both")
        want_sip = mode in ("SIP", "Both")

        for attempt in range(
            1,
            MAX_RETRIES + 1
        ):

            page = await context.new_page()

            try:

                print(
                    f"[{attempt}/{MAX_RETRIES}] "
                    f"{mode} -> {name}"
                )

                await prepare_page(
                    page
                )

                await load_fund_page(
                    page,
                    url
                )

                # ------------------------------------------------
                # AUM
                # ------------------------------------------------

                aum = await extract_aum(
                    page
                )

                # ------------------------------------------------
                # Performance
                # ------------------------------------------------

                performance = await get_performance_section(
                    page
                )

                if performance is None:

                    raise RuntimeError(
                        "Performance section not found"
                    )

                # No "Performance" tab click needed: #performance is
                # already present on the page, showing the Lumpsum
                # table by default. extract_sip() clicks the SIP label
                # itself when needed (see click_sip()).
                lumpsum_returns = empty_returns()
                sip_returns = empty_returns()

                if want_lumpsum:
                    lumpsum_returns = await extract_lumpsum(
                        page,
                        performance
                    )

                if want_sip:
                    sip_returns = await extract_sip(
                        page,
                        performance
                    )

                print(
                    f"RESULT | {mode} | {name} | AUM={aum} | "
                    f"Lumpsum={lumpsum_returns if want_lumpsum else 'skipped'} | "
                    f"SIP={sip_returns if want_sip else 'skipped'}"
                )

                # ------------------------------------------------
                # Success condition
                # ------------------------------------------------

                got_data = (
                    aum is not None
                    or (want_lumpsum and has_any_returns(lumpsum_returns))
                    or (want_sip and has_any_returns(sip_returns))
                )

                if not got_data:

                    raise RuntimeError(
                        "No data extracted"
                    )

                row = {
                    "Fund": name,
                    "AUM (₹ Cr.)": aum,
                }

                if mode == "Both":
                    # Distinct column names so Lumpsum and SIP numbers
                    # never collide when both are requested together.
                    for key in PERIODS:
                        row[f"{key} Lumpsum (%)"] = lumpsum_returns[key]
                    for key in PERIODS:
                        row[f"{key} SIP (%)"] = sip_returns[key]
                elif mode == "Lumpsum":
                    for key in PERIODS:
                        row[f"{key} (%)"] = lumpsum_returns[key]
                else:  # SIP
                    for key in PERIODS:
                        row[f"{key} (%)"] = sip_returns[key]

                row["Source"] = url

                return row

            except Exception as e:

                print(
                    f"ERROR | {mode} | "
                    f"{name} | "
                    f"attempt={attempt}: {e}"
                )

                if attempt < MAX_RETRIES:

                    await asyncio.sleep(attempt)

            finally:

                try:
                    await page.close()
                except Exception:
                    pass

        # --------------------------------------------------------
        # Failed fund
        # --------------------------------------------------------

        failed_row = {
            "Fund": name,
            "AUM (₹ Cr.)": None,
        }

        if mode == "Both":
            for key in PERIODS:
                failed_row[f"{key} Lumpsum (%)"] = None
            for key in PERIODS:
                failed_row[f"{key} SIP (%)"] = None
        else:
            for key in PERIODS:
                failed_row[f"{key} (%)"] = None

        failed_row["Source"] = url

        return failed_row


# ============================================================
# ASYNC SCRAPER
# ============================================================

async def _scrape(
    selected_funds: List[str],
    mode: str
):

    if mode == "Both":
        columns = (
            ["Fund", "AUM (₹ Cr.)"]
            + [f"{k} Lumpsum (%)" for k in PERIODS]
            + [f"{k} SIP (%)" for k in PERIODS]
            + ["Source"]
        )
    else:
        columns = (
            ["Fund", "AUM (₹ Cr.)"]
            + [f"{k} (%)" for k in PERIODS]
            + ["Source"]
        )

    valid_funds = [
        fund
        for fund in selected_funds
        if fund in FUNDS
    ]

    if not valid_funds:

        return pd.DataFrame(
            columns=columns
        )

    # Ensure virtual display (no-op when HEADLESS=True).
    start_virtual_display()

    # Ensure browser.
    ensure_playwright_browser()

    async with async_playwright() as playwright:

        browser = await create_browser(
            playwright
        )

        context = await create_context(
            browser
        )

        try:

            semaphore = asyncio.Semaphore(
                MAX_CONCURRENCY
            )

            tasks = [
                scrape_one(
                    context,
                    semaphore,
                    fund,
                    mode
                )
                for fund in valid_funds
            ]

            results = await asyncio.gather(
                *tasks
            )

        finally:

            try:
                await context.close()
            except Exception:
                pass

            try:
                await browser.close()
            except Exception:
                pass

    df = pd.DataFrame(
        results
    ).reindex(
        columns=columns
    )

    numeric_columns = [
        column for column in columns
        if column not in ("Fund", "Source")
    ]

    for column in numeric_columns:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

    return df


# ============================================================
# PUBLIC SCRAPER
# ============================================================

def scrape_funds(
    selected_funds,
    mode="Lumpsum"
):
    """
    mode:
      - "Lumpsum": only Lumpsum Annualised(%) columns.
      - "SIP": only SIP Annualised(%) columns.
      - "Both": Lumpsum AND SIP columns, fetched from a single page
        load per fund (faster than calling this twice).
    """

    if mode not in {
        "Lumpsum",
        "SIP",
        "Both",
    }:

        raise ValueError(
            "mode must be Lumpsum, SIP, or Both"
        )

    # Streamlit runs this synchronously. Multi-fund scraping is concurrent.
    return asyncio.run(
        _scrape(
            selected_funds,
            mode
        )
    )


# ============================================================
# EXCEL
# ============================================================

def save_to_excel(
    df,
    output_path=OUTPUT_FILE
):
    """
    FIX: this used to hardcode a fixed 6-column layout, so any extra
    columns (e.g. the "Both" mode's separate Lumpsum/SIP columns) were
    silently dropped by the reindex(). It now keeps whatever columns
    the DataFrame actually has (minus "Source"), and uses
    get_column_letter() instead of chr(64 + index) for column widths.
    """

    output = df.drop(
        columns=["Source"],
        errors="ignore"
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        output.to_excel(
            writer,
            sheet_name="Mutual Funds",
            index=False
        )

    workbook = load_workbook(
        output_path
    )

    worksheet = workbook[
        "Mutual Funds"
    ]

    fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    for cell in worksheet[1]:

        cell.fill = fill

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    for row in worksheet.iter_rows(
        min_row=2
    ):

        for cell in row:

            cell.alignment = Alignment(
                vertical="center"
            )

    for index, column_name in enumerate(
        output.columns,
        start=1
    ):

        width = 48 if column_name == "Fund" else 16

        worksheet.column_dimensions[
            get_column_letter(index)
        ].width = width

    workbook.save(
        output_path
    )

    return str(output_path)


# ============================================================
# LOCAL TEST
# ============================================================

if __name__ == "__main__":

    test_fund = [
        "Bandhan Small Cap Fund Regular Growth"
    ]

    print(
        "\n========== BOTH (Lumpsum + SIP in one page load) ==========\n"
    )

    both = scrape_funds(
        test_fund,
        "Both"
    )

    print(
        both.to_string(
            index=False
        )
    )