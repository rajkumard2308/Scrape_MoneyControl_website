import asyncio
import os
import re
import shutil
import subprocess
import sys
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# SETTINGS
# ============================================================

OUTPUT_FILE = "mutual_funds.xlsx"

# IMPORTANT:
# We deliberately use headed Chromium through Xvfb.
# This behaves much closer to the working local version.
HEADLESS = False

# Moneycontrol + Streamlit Cloud
MAX_CONCURRENCY = 3

PAGE_TIMEOUT = 60000
DATA_TIMEOUT = 30000
ELEMENT_TIMEOUT = 15000

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
    Start Xvfb when running on Linux/Streamlit Cloud.

    This allows Playwright to run in headed mode without
    a physical display.
    """

    global _XVFB_PROCESS

    # Windows/macOS local machine already has a display.
    if os.name != "posix":
        return

    # If DISPLAY already exists, don't start another one.
    if os.environ.get("DISPLAY"):
        return

    xvfb = shutil.which("Xvfb")

    if not xvfb:
        print(
            "WARNING: Xvfb not found. "
            "Playwright headed mode may fail."
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


# Start display as soon as this module loads.
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

    browser = await playwright.chromium.launch(

        # IMPORTANT:
        # Headed mode is intentional.
        # Xvfb provides the virtual display on Cloud.
        headless=False,

        args=[
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
        ],
    )

    return browser


async def create_context(browser):

    context = await browser.new_context(

        viewport={
            "width": 1440,
            "height": 900,
        },

        screen={
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

    # Moneycontrol frontend rendering.
    await page.wait_for_timeout(
        2500
    )

    try:

        await page.wait_for_load_state(
            "networkidle",
            timeout=10000
        )

    except PlaywrightTimeoutError:
        pass

    await page.wait_for_timeout(
        1500
    )


# ============================================================
# PERFORMANCE TAB
# ============================================================

async def activate_performance(page):

    candidates = [

        page.get_by_role(
            "button",
            name=re.compile(
                r"^Performance$",
                re.I
            )
        ).first,

        page.get_by_text(
            "Performance",
            exact=True
        ).first,

        page.locator(
            "#performance"
        ).first,
    ]

    for candidate in candidates:

        try:

            if await candidate.count() == 0:
                continue

            if await candidate.is_visible():

                await candidate.scroll_into_view_if_needed(
                    timeout=3000
                )

                try:

                    await candidate.click(
                        timeout=5000
                    )

                except Exception:

                    await candidate.click(
                        timeout=5000,
                        force=True
                    )

                await page.wait_for_timeout(
                    1000
                )

                return

        except Exception:
            continue


# ============================================================
# AUM
# ============================================================

async def extract_aum(page):

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

    performance = page.locator(
        "#performance"
    ).first

    candidates = [

        performance.get_by_text(
            "SIP",
            exact=True
        ).first,

        performance.get_by_role(
            "radio",
            name=re.compile(
                r"^SIP$",
                re.I
            )
        ).first,

        performance.locator(
            "label"
        ).filter(
            has_text=re.compile(
                r"^\s*SIP\s*$",
                re.I
            )
        ).first,

        performance.locator(
            'input[type="radio"]'
        ).nth(1),

        page.get_by_text(
            "SIP",
            exact=True
        ).first,
    ]

    for candidate in candidates:

        try:

            if await candidate.count() == 0:
                continue

            if not await candidate.is_visible():
                continue

            await candidate.scroll_into_view_if_needed(
                timeout=3000
            )

            try:

                await candidate.click(
                    timeout=5000
                )

            except Exception:

                await candidate.click(
                    timeout=5000,
                    force=True
                )

            # Wait for Moneycontrol JS.
            await page.wait_for_timeout(
                2000
            )

            return True

        except Exception:
            continue

    return False


# ============================================================
# TABLE READER
# ============================================================

async def read_table(table):

    rows = table.locator(
        "tr"
    )

    count = await rows.count()

    result = []

    for i in range(count):

        cells = rows.nth(i).locator(
            "th, td"
        )

        cell_count = await cells.count()

        if cell_count == 0:
            continue

        row = []

        for j in range(cell_count):

            try:

                value = clean_text(
                    await cells.nth(j).inner_text(
                        timeout=1500
                    )
                )

            except Exception:

                value = ""

            row.append(value)

        if row:
            result.append(row)

    return result


# ============================================================
# TABLE PARSER
# ============================================================

def find_annualised_column(rows):

    # Search first 5 rows.
    for row in rows[:5]:

        for index, value in enumerate(row):

            text = clean_text(
                value
            ).lower()

            if (
                "annualised" in text
                or "annualized" in text
            ):
                return index

    return None


async def parse_table(table, mode):

    rows = await read_table(
        table
    )

    result = empty_returns()

    if not rows:
        return result

    annualised_index = find_annualised_column(
        rows
    )

    # --------------------------------------------------------
    # Header based extraction
    # --------------------------------------------------------

    if annualised_index is not None:

        for row in rows:

            if not row:
                continue

            key = period_key(
                row[0]
            )

            if not key:
                continue

            if annualised_index >= len(row):
                continue

            value = parse_number(
                row[annualised_index]
            )

            if value is not None:

                result[key] = value

    # --------------------------------------------------------
    # Known Moneycontrol structure fallback
    # --------------------------------------------------------

    for row in rows:

        if not row:
            continue

        key = period_key(
            row[0]
        )

        if not key:
            continue

        if result[key] is not None:
            continue

        if mode == "Lumpsum":

            # Period | Absolute | Annualised
            if len(row) >= 3:

                value = parse_number(
                    row[2]
                )

                if value is not None:
                    result[key] = value

        else:

            # Period | Start | Invested | Latest |
            # Absolute | Annualised

            if len(row) >= 2:

                value = parse_number(
                    row[-1]
                )

                if value is not None:
                    result[key] = value

    return result


# ============================================================
# TEXT PARSER
# ============================================================

def parse_returns_from_text(
    text,
    mode
):

    text = clean_text(
        text
    )

    result = empty_returns()

    if not text:
        return result

    text = re.sub(
        r"Annualized",
        "Annualised",
        text,
        flags=re.I
    )

    period_pattern = (
        r"(1\s*(?:Year|Years|Y|Yr|Yrs)|"
        r"2\s*(?:Year|Years|Y|Yr|Yrs)|"
        r"3\s*(?:Year|Years|Y|Yr|Yrs)|"
        r"5\s*(?:Year|Years|Y|Yr|Yrs))"
    )

    matches = list(
        re.finditer(
            period_pattern,
            text,
            re.I
        )
    )

    for index, match in enumerate(matches):

        key = period_key(
            match.group(1)
        )

        if not key:
            continue

        start = match.end()

        if index + 1 < len(matches):

            end = matches[
                index + 1
            ].start()

        else:

            end = min(
                len(text),
                start + 300
            )

        section = text[
            start:end
        ]

        numbers = re.findall(
            r"-?\d[\d,]*(?:\.\d+)?%?",
            section
        )

        values = []

        for item in numbers:

            value = parse_number(
                item
            )

            if value is not None:

                values.append(
                    value
                )

        if values:

            # Annualised is the last return value
            # in the Moneycontrol return row.
            result[key] = values[-1]

    return result


# ============================================================
# PERFORMANCE TEXT
# ============================================================

async def performance_text(
    page,
    performance
):

    try:

        text = await performance.inner_text(
            timeout=5000
        )

        return clean_text(
            text
        )

    except Exception:
        pass

    try:

        text = await page.locator(
            "body"
        ).inner_text(
            timeout=5000
        )

        return clean_text(
            text
        )

    except Exception:

        return ""


# ============================================================
# LUMPSUM
# ============================================================

async def extract_lumpsum(
    page,
    performance
):

    result = empty_returns()

    await activate_performance(
        page
    )

    # Wait for annualised content.
    try:

        await page.wait_for_function(
            """
            () => {
                const el =
                    document.querySelector("#performance");

                if (!el) return false;

                const text =
                    el.innerText || "";

                return (
                    text.includes("Annualised") ||
                    text.includes("Annualized") ||
                    text.includes("1 Year")
                );
            }
            """,
            timeout=DATA_TIMEOUT
        )

    except Exception:
        pass

    await page.wait_for_timeout(
        1000
    )

    # --------------------------------------------------------
    # Parse tables
    # --------------------------------------------------------

    tables = performance.locator(
        "table"
    )

    count = await tables.count()

    for i in range(count):

        try:

            table_text = clean_text(
                await tables.nth(i).inner_text(
                    timeout=3000
                )
            )

            if not table_text:
                continue

            if (
                "Annualised" not in table_text
                and "Annualized" not in table_text
                and "1 Year" not in table_text
            ):
                continue

            parsed = await parse_table(
                tables.nth(i),
                "Lumpsum"
            )

            for key in PERIODS:

                if parsed[key] is not None:

                    result[key] = (
                        parsed[key]
                    )

            if has_all_returns(result):

                return result

        except Exception:
            continue

    # --------------------------------------------------------
    # Text fallback
    # --------------------------------------------------------

    text = await performance_text(
        page,
        performance
    )

    parsed = parse_returns_from_text(
        text,
        "Lumpsum"
    )

    for key in PERIODS:

        if result[key] is None:

            result[key] = parsed[key]

    return result


# ============================================================
# SIP
# ============================================================

async def extract_sip(
    page,
    performance
):

    result = empty_returns()

    clicked = await click_sip(
        page
    )

    if not clicked:

        print(
            "WARNING: SIP button not found"
        )

        return result

    # Wait for SIP data.
    try:

        await page.wait_for_function(
            """
            () => {
                const el =
                    document.querySelector("#performance");

                if (!el) return false;

                const text =
                    el.innerText || "";

                return (
                    text.includes("Invested") ||
                    text.includes("SIP Returns") ||
                    text.includes("Annualised") ||
                    text.includes("Annualized")
                );
            }
            """,
            timeout=DATA_TIMEOUT
        )

    except Exception:
        pass

    await page.wait_for_timeout(
        1200
    )

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    tables = performance.locator(
        "table"
    )

    count = await tables.count()

    for i in range(count):

        try:

            table_text = clean_text(
                await tables.nth(i).inner_text(
                    timeout=3000
                )
            )

            if not table_text:
                continue

            if (
                "Invested" not in table_text
                and "SIP Returns" not in table_text
                and "Annualised" not in table_text
                and "Annualized" not in table_text
            ):
                continue

            parsed = await parse_table(
                tables.nth(i),
                "SIP"
            )

            for key in PERIODS:

                if parsed[key] is not None:

                    result[key] = (
                        parsed[key]
                    )

            if has_all_returns(result):

                return result

        except Exception:
            continue

    # --------------------------------------------------------
    # Text fallback
    # --------------------------------------------------------

    text = await performance_text(
        page,
        performance
    )

    parsed = parse_returns_from_text(
        text,
        "SIP"
    )

    for key in PERIODS:

        if result[key] is None:

            result[key] = parsed[key]

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

    async with semaphore:

        url = FUNDS[name]["url"]

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

                if mode == "SIP":

                    returns = await extract_sip(
                        page,
                        performance
                    )

                else:

                    returns = await extract_lumpsum(
                        page,
                        performance
                    )

                print(
                    f"RESULT | {mode} | {name} | "
                    f"AUM={aum} | "
                    f"1Y={returns['1Y']} | "
                    f"2Y={returns['2Y']} | "
                    f"3Y={returns['3Y']} | "
                    f"5Y={returns['5Y']}"
                )

                # ------------------------------------------------
                # Success condition
                # ------------------------------------------------

                if (
                    aum is None
                    and not has_any_returns(
                        returns
                    )
                ):

                    raise RuntimeError(
                        "No data extracted"
                    )

                return {
                    "Fund": name,

                    "AUM (₹ Cr.)": aum,

                    "1Y (%)":
                        returns["1Y"],

                    "2Y (%)":
                        returns["2Y"],

                    "3Y (%)":
                        returns["3Y"],

                    "5Y (%)":
                        returns["5Y"],

                    "Source": url,
                }

            except Exception as e:

                print(
                    f"ERROR | {mode} | "
                    f"{name} | "
                    f"attempt={attempt}: {e}"
                )

                if attempt < MAX_RETRIES:

                    await asyncio.sleep(
                        2 * attempt
                    )

            finally:

                try:
                    await page.close()
                except Exception:
                    pass

        # --------------------------------------------------------
        # Failed fund
        # --------------------------------------------------------

        return {
            "Fund": name,
            "AUM (₹ Cr.)": None,
            "1Y (%)": None,
            "2Y (%)": None,
            "3Y (%)": None,
            "5Y (%)": None,
            "Source": url,
        }


# ============================================================
# ASYNC SCRAPER
# ============================================================

async def _scrape(
    selected_funds: List[str],
    mode: str
):

    columns = [
        "Fund",
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
        "Source",
    ]

    valid_funds = [
        fund
        for fund in selected_funds
        if fund in FUNDS
    ]

    if not valid_funds:

        return pd.DataFrame(
            columns=columns
        )

    # Ensure virtual display.
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

    for column in [
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]:

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

    if mode not in {
        "Lumpsum",
        "SIP",
    }:

        raise ValueError(
            "mode must be Lumpsum or SIP"
        )

    # Streamlit runs this synchronously.
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

    columns = [
        "Fund",
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]

    output = (
        df
        .drop(
            columns=["Source"],
            errors="ignore"
        )
        .reindex(
            columns=[
                column
                for column in columns
                if column in df.columns
            ]
        )
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

    widths = [
        48,
        18,
        12,
        12,
        12,
        12,
    ]

    for index, width in enumerate(
        widths,
        start=1
    ):

        worksheet.column_dimensions[
            chr(64 + index)
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
        "\n========== LUMPSUM ==========\n"
    )

    lumpsum = scrape_funds(
        test_fund,
        "Lumpsum"
    )

    print(
        lumpsum.to_string(
            index=False
        )
    )

    print(
        "\n========== SIP ==========\n"
    )

    sip = scrape_funds(
        test_fund,
        "SIP"
    )

    print(
        sip.to_string(
            index=False
        )
    )