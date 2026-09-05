import asyncio
import re
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
# Keep True for Streamlit Cloud.
HEADLESS = True

# 3 is a good balance for Moneycontrol.
# Increase to 4 only after confirming stability.
MAX_CONCURRENCY = 3

PAGE_TIMEOUT = 60000
PERFORMANCE_TIMEOUT = 25000
ELEMENT_TIMEOUT = 15000

# Number of attempts for an individual fund.
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
# HELPERS
# ============================================================

def clean_text(value) -> str:
    return re.sub(
        r"\s+",
        " ",
        str(value or "").replace("\xa0", " ")
    ).strip()


def parse_number(value) -> Optional[float]:
    value = clean_text(value)

    if not value:
        return None

    if value in {
        "--",
        "-",
        "N/A",
        "NA",
        "Nil",
        "n.a.",
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
            match.group(0).replace(",", "")
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
    value = clean_text(value).lower()

    if re.search(r"\b1\s*year\b", value):
        return "1Y"

    if re.search(r"\b2\s*years?\b", value):
        return "2Y"

    if re.search(r"\b3\s*years?\b", value):
        return "3Y"

    if re.search(r"\b5\s*years?\b", value):
        return "5Y"

    return None


def has_all_returns(data: Dict[str, Optional[float]]) -> bool:
    return all(
        data.get(key) is not None
        for key in ("1Y", "2Y", "3Y", "5Y")
    )


# ============================================================
# BROWSER
# ============================================================

async def create_browser(playwright):
    """
    Launch Debian's system Chromium headless shell.
    This is used on Streamlit Community Cloud.
    """

    browser = await playwright.chromium.launch(
        headless=True,
        executable_path="/usr/bin/chromium-headless-shell",
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
            "--disable-features=Translate,BackForwardCache",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-notifications",
            "--disable-popup-blocking",
            "--window-size=1440,900",
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
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/131.0.0.0 Safari/537.36"
        ),

        extra_http_headers={
            "Accept-Language": "en-IN,en;q=0.9",
            "Upgrade-Insecure-Requests": "1",
        },

        service_workers="block",
    )

    return context


# ============================================================
# PAGE PREPARATION
# ============================================================

async def prepare_page(page):

    # Don't block images/fonts here.
    # Moneycontrol can behave differently if these are blocked.
    #
    # We only abort obvious tracking/advertising requests.
    async def route_handler(route):

        request = route.request
        url = request.url.lower()

        # Keep the main site resources.
        # Avoid aggressively blocking because the site is dynamic.
        if any(
            x in url
            for x in [
                "doubleclick.net",
                "googlesyndication.com",
                "googleadservices.com",
                "facebook.com/tr",
            ]
        ):
            try:
                await route.abort()
            except Exception:
                pass
            return

        try:
            await route.continue_()
        except Exception:
            pass

    await page.route(
        "**/*",
        route_handler
    )

    # Prevent dialogs from stopping the scraper.
    page.on(
        "dialog",
        lambda dialog: asyncio.create_task(
            dialog.dismiss()
        )
    )


# ============================================================
# PAGE LOAD
# ============================================================

async def load_fund_page(page, url):

    response = await page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT,
    )

    if response:
        status = response.status

        print(
            f"HTTP {status}: {url}"
        )

        if status >= 400:
            raise RuntimeError(
                f"HTTP {status}"
            )

    # Let Moneycontrol's JS render.
    await page.wait_for_timeout(1800)

    # DOM may continue loading after domcontentloaded.
    try:
        await page.wait_for_load_state(
            "networkidle",
            timeout=8000,
        )
    except PlaywrightTimeoutError:
        # Networkidle isn't required.
        pass

    await page.wait_for_timeout(700)


# ============================================================
# AUM
# ============================================================

async def extract_aum(page):

    selectors = [
        '//*[@id="overview"]/div[1]/ul/li[1]',

        "#overview",

        'text=/AUM\\s*\\(Crs?\\.?\\)/i',
    ]

    for selector in selectors:

        try:

            if selector.startswith("text="):
                locator = page.get_by_text(
                    re.compile(
                        r"AUM\s*\(Crs?\.?\)",
                        re.I
                    )
                ).first
            elif selector.startswith("//"):
                locator = page.locator(
                    f"xpath={selector}"
                ).first
            else:
                locator = page.locator(
                    selector
                ).first

            if await locator.count() == 0:
                continue

            await locator.wait_for(
                state="attached",
                timeout=ELEMENT_TIMEOUT
            )

            text = clean_text(
                await locator.inner_text(
                    timeout=5000
                )
            )

            # Example:
            # AUM (Crs.) 12,345.67
            match = re.search(
                r"AUM\s*\(Crs?\.?\)\s*"
                r"([\d,]+(?:\.\d+)?)",
                text,
                re.I
            )

            if match:
                return parse_number(
                    match.group(1)
                )

            # Fallback
            numbers = re.findall(
                r"\d[\d,]*(?:\.\d+)?",
                text
            )

            if numbers:
                return parse_number(
                    numbers[-1]
                )

        except Exception:
            continue

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
            timeout=PERFORMANCE_TIMEOUT
        )

        await performance.scroll_into_view_if_needed(
            timeout=5000
        )

        return performance

    except Exception as e:

        print(
            "Performance section error:",
            e
        )

        return None


# ============================================================
# SIP BUTTON
# ============================================================

async def click_sip(page):

    performance = page.locator(
        "#performance"
    ).first

    candidates = [

        # Text
        performance.get_by_text(
            "SIP",
            exact=True
        ).first,

        # Label
        performance.locator(
            "label"
        ).filter(
            has_text=re.compile(
                r"^\s*SIP\s*$",
                re.I
            )
        ).first,

        # Radio
        performance.locator(
            'input[type="radio"]'
        ).nth(1),

        # Button
        performance.locator(
            "button"
        ).filter(
            has_text=re.compile(
                r"^\s*SIP\s*$",
                re.I
            )
        ).first,

        # Generic element
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

            # Click with force only if normal click fails.
            try:

                await candidate.click(
                    timeout=5000
                )

            except Exception:

                await candidate.click(
                    timeout=5000,
                    force=True
                )

            # Give JS time to switch table.
            await page.wait_for_timeout(
                1000
            )

            return True

        except Exception:
            continue

    return False


# ============================================================
# TABLE EXTRACTION
# ============================================================

async def read_table_rows(table):

    rows = table.locator("tr")

    row_count = await rows.count()

    all_rows = []

    for i in range(row_count):

        cells = rows.nth(i).locator(
            "th, td"
        )

        cell_count = await cells.count()

        if cell_count == 0:
            continue

        values = []

        for j in range(cell_count):

            try:

                text = clean_text(
                    await cells.nth(j).inner_text(
                        timeout=1500
                    )
                )

            except Exception:

                text = ""

            values.append(text)

        if values:
            all_rows.append(values)

    return all_rows


def find_annualised_column(rows):

    if not rows:
        return None

    # Search first 3 rows because headers can span rows.
    for row in rows[:3]:

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

    rows = await read_table_rows(
        table
    )

    if not rows:
        return empty_returns()

    annualised_index = find_annualised_column(
        rows
    )

    result = empty_returns()

    # --------------------------------------------------------
    # Preferred method:
    # Find Annualised column from header.
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

            if annualised_index < len(row):

                value = parse_number(
                    row[annualised_index]
                )

                if value is not None:
                    result[key] = value

    # --------------------------------------------------------
    # Fallback based on known Moneycontrol structures.
    # --------------------------------------------------------

    if not has_all_returns(result):

        for row in rows:

            if not row:
                continue

            key = period_key(
                row[0]
            )

            if not key:
                continue

            # Lumpsum:
            # Period | Absolute | Annualised
            if mode == "Lumpsum":

                if len(row) >= 3:

                    value = parse_number(
                        row[2]
                    )

                    if value is not None:
                        result[key] = value

            # SIP:
            # Period | Start | Invested | Latest |
            # Absolute | Annualised
            else:

                if len(row) >= 2:

                    value = parse_number(
                        row[-1]
                    )

                    if value is not None:
                        result[key] = value

    return result


# ============================================================
# PERFORMANCE EXTRACTION
# ============================================================

async def extract_lumpsum(page, performance):

    result = empty_returns()

    # Find every table in performance section.
    tables = performance.locator(
        "table"
    )

    count = await tables.count()

    for i in range(count):

        try:

            parsed = await parse_table(
                tables.nth(i),
                "Lumpsum"
            )

            for key in result:

                if parsed[key] is not None:
                    result[key] = parsed[key]

            if has_all_returns(result):
                return result

        except Exception:
            continue

    # --------------------------------------------------------
    # Text fallback
    # --------------------------------------------------------

    try:

        text = clean_text(
            await performance.inner_text(
                timeout=5000
            )
        )

        patterns = {

            "1Y": (
                r"1\s*Year\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)"
            ),

            "2Y": (
                r"2\s*Years?\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)"
            ),

            "3Y": (
                r"3\s*Years?\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)"
            ),

            "5Y": (
                r"5\s*Years?\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)\s+"
                r"(-?[\d,]+(?:\.\d+)?|--)"
            ),
        }

        for key, pattern in patterns.items():

            if result[key] is not None:
                continue

            match = re.search(
                pattern,
                text,
                re.I
            )

            if match:

                result[key] = parse_number(
                    match.group(2)
                )

    except Exception:
        pass

    return result


async def extract_sip(page, performance):

    result = empty_returns()

    clicked = await click_sip(
        page
    )

    if not clicked:

        print(
            "SIP button not found"
        )

        return result

    # --------------------------------------------------------
    # Wait until SIP content appears.
    # --------------------------------------------------------

    try:

        await page.wait_for_function(
            """
            () => {
                const el =
                    document.querySelector("#performance");

                if (!el) {
                    return false;
                }

                const text =
                    el.innerText || "";

                return (
                    text.includes("Invested") ||
                    text.includes("SIP Returns")
                );
            }
            """,
            timeout=12000
        )

    except Exception:
        pass

    await page.wait_for_timeout(
        800
    )

    # --------------------------------------------------------
    # Parse SIP tables.
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

            # Prefer SIP table.
            if (
                "Invested" not in table_text
                and "Annualised" not in table_text
            ):
                continue

            parsed = await parse_table(
                tables.nth(i),
                "SIP"
            )

            for key in result:

                if parsed[key] is not None:
                    result[key] = parsed[key]

            if has_all_returns(result):
                return result

        except Exception:
            continue

    # --------------------------------------------------------
    # SIP text fallback
    # --------------------------------------------------------

    try:

        text = clean_text(
            await performance.inner_text(
                timeout=5000
            )
        )

        # Match each period and take the last numeric
        # value before the next period.
        for key, period_pattern in {
            "1Y": r"1\s*Year",
            "2Y": r"2\s*Years?",
            "3Y": r"3\s*Years?",
            "5Y": r"5\s*Years?",
        }.items():

            if result[key] is not None:
                continue

            match = re.search(
                period_pattern +
                r"(.*?)(?="
                r"\d\s*Years?"
                r"|$)",
                text,
                re.I
            )

            if not match:
                continue

            section = match.group(1)

            numbers = re.findall(
                r"-?\d[\d,]*(?:\.\d+)?",
                section
            )

            if numbers:

                result[key] = parse_number(
                    numbers[-1]
                )

    except Exception:
        pass

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
                        "Performance section not loaded"
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
                    f"AUM={aum} | {returns}"
                )

                # If nothing was obtained, retry.
                if (
                    aum is None
                    and not any(
                        value is not None
                        for value in returns.values()
                    )
                ):

                    raise RuntimeError(
                        "No data extracted"
                    )

                return {
                    "Fund": name,
                    "AUM (₹ Cr.)": aum,
                    "1Y (%)": returns["1Y"],
                    "2Y (%)": returns["2Y"],
                    "3Y (%)": returns["3Y"],
                    "5Y (%)": returns["5Y"],
                    "Source": url,
                }

            except Exception as e:

                print(
                    f"ERROR | {mode} | {name} | "
                    f"attempt={attempt}: {e}"
                )

                if attempt < MAX_RETRIES:

                    # Small retry delay.
                    await asyncio.sleep(
                        1.0 * attempt
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
# MAIN ASYNC SCRAPER
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

            # IMPORTANT:
            # Same browser + same context.
            # Pages are opened concurrently.
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
# PUBLIC SCRAPER FUNCTION
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

    try:

        return asyncio.run(
            _scrape(
                selected_funds,
                mode
            )
        )

    except RuntimeError as e:

        # Safety for environments where an event loop
        # already exists.
        if "asyncio.run()" not in str(e):
            raise

        loop = asyncio.new_event_loop()

        try:

            asyncio.set_event_loop(
                loop
            )

            return loop.run_until_complete(
                _scrape(
                    selected_funds,
                    mode
                )
            )

        finally:

            loop.close()

            asyncio.set_event_loop(
                None
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

    print("\n========== LUMPSUM ==========\n")

    lumpsum = scrape_funds(
        test_fund,
        "Lumpsum"
    )

    print(
        lumpsum.to_string(
            index=False
        )
    )

    print("\n========== SIP ==========\n")

    sip = scrape_funds(
        test_fund,
        "SIP"
    )

    print(
        sip.to_string(
            index=False
        )
    )