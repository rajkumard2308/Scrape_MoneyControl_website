import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ============================================================
# CONFIG
# ============================================================

OUTPUT_FILE = "mutual_funds.xlsx"

# Always use headless mode for Streamlit Cloud.
HEADLESS = True

PAGE_TIMEOUT = 45_000
ELEMENT_TIMEOUT = 20_000
REQUEST_DELAY = 0.5
MAX_RETRIES = 2


# ============================================================
# MONEYCONTROL TARGETS
# ============================================================
# We do NOT scroll the whole page.
# We only target these two Moneycontrol sections.
# ============================================================

AUM_XPATH = '//*[@id="overview"]/div[1]/ul/li[1]'

PERFORMANCE_XPATH = '//*[@id="performance"]/div/div[3]/div[1]'


# ============================================================
# MONEYCONTROL FUND URLS
# ============================================================

FUNDS = {
    "Bandhan Small Cap Fund Regular Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-small-cap-fund-regular-plan-growth/MAG2106"
    },

    "Bandhan Large & Mid Cap Fund - Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-large-mid-cap-fund-regular-plan-growth/MAG091"
    },

    "Bandhan Innovation Fund Regular Plan Growth": {
        "url": "https://www.moneycontrol.com/mutual-funds/nav/bandhan-innovation-fund-regular-plan-growth/MAGA044"
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
# TEXT / NUMBER HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def parse_number(value):
    if value is None:
        return None

    value = clean_text(value)

    if value.upper() in {"", "--", "-", "NA", "N/A", "NULL"}:
        return None

    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", value)

    if not match:
        return None

    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


# ============================================================
# BROWSER
# ============================================================

def launch_browser(playwright):
    args = [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--disable-notifications",
        "--disable-popup-blocking",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
    ]

    # Streamlit Cloud / Linux
    if sys.platform.startswith("linux"):
        possible_browsers = [
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
        ]

        for executable in possible_browsers:
            if os.path.exists(executable):
                return playwright.chromium.launch(
                    executable_path=executable,
                    headless=True,
                    args=args,
                )

        # If no system Chromium exists, use Playwright's installed browser.
        return playwright.chromium.launch(
            headless=True,
            args=args,
        )

    # Windows
    return playwright.chromium.launch(
        headless=HEADLESS,
        channel="chromium",
        args=args,
    )


def create_context(browser):
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/139.0.0.0 Safari/537.36"
        ),
        locale="en-IN",
        timezone_id="Asia/Kolkata",
        java_script_enabled=True,
    )

    context.add_init_script(
        """
        Object.defineProperty(
            navigator,
            'webdriver',
            { get: () => undefined }
        );
        """
    )

    # Block only heavy resources. JavaScript remains enabled.
    def route_handler(route):
        if route.request.resource_type in {"image", "media", "font"}:
            route.abort()
        else:
            route.continue_()

    context.route("**/*", route_handler)

    return context


# ============================================================
# AUM
# ============================================================

def extract_aum(page):
    locator = page.locator(f"xpath={AUM_XPATH}")

    try:
        locator.wait_for(
            state="attached",
            timeout=ELEMENT_TIMEOUT,
        )

        # Targeted scroll only to this card.
        # We never scroll through 25/50/75/100% of the page.
        try:
            locator.scroll_into_view_if_needed(timeout=5_000)
        except Exception:
            pass

        text = clean_text(
            locator.inner_text(timeout=5_000)
        )

        print(f"AUM card: {text}")

        patterns = [
            r"AUM\s*\(Crs\.\)\s*([\d,]+(?:\.\d+)?)",
            r"AUM\s*\(Cr\.\)\s*([\d,]+(?:\.\d+)?)",
            r"AUM\s*\(Crs\)\s*([\d,]+(?:\.\d+)?)",
            r"AUM\s*([\d,]+(?:\.\d+)?)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)

            if match:
                value = parse_number(match.group(1))
                if value is not None:
                    return value

        # Fallback: only within the AUM card.
        numbers = re.findall(r"\d[\d,]*(?:\.\d+)?", text)

        if numbers:
            return parse_number(numbers[-1])

    except Exception as exc:
        print(f"AUM extraction error: {exc}")

    return None


# ============================================================
# PERFORMANCE
# ============================================================

def extract_performance(page):
    result = {
        "1Y": None,
        "2Y": None,
        "3Y": None,
        "5Y": None,
    }

    locator = page.locator(f"xpath={PERFORMANCE_XPATH}")

    try:
        locator.wait_for(
            state="attached",
            timeout=ELEMENT_TIMEOUT,
        )

        # Targeted scroll to the performance section only.
        try:
            locator.scroll_into_view_if_needed(timeout=5_000)
        except Exception:
            pass

    except PlaywrightTimeoutError:
        print("Performance section not found.")
        return result
    except Exception as exc:
        print(f"Performance locator error: {exc}")
        return result

    # --------------------------------------------------------
    # Preferred: read table rows inside the target section.
    # --------------------------------------------------------

    try:
        rows = locator.locator("tr")
        row_count = rows.count()

        print(f"Performance rows found: {row_count}")

        for i in range(row_count):
            row = rows.nth(i)
            cells = row.locator("th, td")
            cell_count = cells.count()

            if cell_count < 2:
                continue

            values = []

            for j in range(cell_count):
                try:
                    values.append(
                        clean_text(
                            cells.nth(j).inner_text(timeout=2_000)
                        )
                    )
                except Exception:
                    values.append("")

            if not values:
                continue

            period = values[0].lower()

            # Moneycontrol:
            # Period | Absolute(%) | Annualised(%)
            #
            # We want Annualised = values[2].
            annualised = (
                parse_number(values[2])
                if len(values) >= 3
                else None
            )

            if re.search(r"\b1\s*Year\b", period, re.I):
                result["1Y"] = annualised

            elif re.search(r"\b2\s*Years?\b", period, re.I):
                result["2Y"] = annualised

            elif re.search(r"\b3\s*Years?\b", period, re.I):
                result["3Y"] = annualised

            elif re.search(r"\b5\s*Years?\b", period, re.I):
                result["5Y"] = annualised

    except Exception as exc:
        print(f"Performance table extraction error: {exc}")

    # --------------------------------------------------------
    # Fallback: read only the target performance card.
    # Never read page.locator("body").
    # --------------------------------------------------------

    if any(value is None for value in result.values()):
        try:
            text = locator.inner_text(timeout=5_000)

            # Capture a period followed by two numeric values.
            # The second numeric value is Annualised.
            patterns = {
                "1Y": r"1\s*Year\s+(-?\d[\d,]*(?:\.\d+)?|--)\s+(-?\d[\d,]*(?:\.\d+)?|--)",
                "2Y": r"2\s*Years?\s+(-?\d[\d,]*(?:\.\d+)?|--)\s+(-?\d[\d,]*(?:\.\d+)?|--)",
                "3Y": r"3\s*Years?\s+(-?\d[\d,]*(?:\.\d+)?|--)\s+(-?\d[\d,]*(?:\.\d+)?|--)",
                "5Y": r"5\s*Years?\s+(-?\d[\d,]*(?:\.\d+)?|--)\s+(-?\d[\d,]*(?:\.\d+)?|--)",
            }

            for key, pattern in patterns.items():
                if result[key] is not None:
                    continue

                match = re.search(
                    pattern,
                    text,
                    re.IGNORECASE,
                )

                if match:
                    result[key] = parse_number(match.group(2))

        except Exception as exc:
            print(f"Performance fallback error: {exc}")

    print(
        "Returns: "
        f"1Y={result['1Y']} | "
        f"2Y={result['2Y']} | "
        f"3Y={result['3Y']} | "
        f"5Y={result['5Y']}"
    )

    return result


# ============================================================
# SCRAPE ONE FUND
# ============================================================

def scrape_fund(page, fund_name, fund_url):
    print()
    print("=" * 70)
    print(f"SCRAPING: {fund_name}")
    print("=" * 70)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"Attempt {attempt}/{MAX_RETRIES}")
            print(f"URL: {fund_url}")

            page.goto(
                fund_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT,
            )

            # No fixed 4-second wait.
            # Wait for the actual target elements instead.
            aum = extract_aum(page)
            performance = extract_performance(page)

            result = {
                "Fund": fund_name,
                "AUM (₹ Cr.)": aum,
                "1Y (%)": performance["1Y"],
                "2Y (%)": performance["2Y"],
                "3Y (%)": performance["3Y"],
                "5Y (%)": performance["5Y"],
                "Source": fund_url,
            }

            has_data = (
                aum is not None
                or any(
                    value is not None
                    for value in performance.values()
                )
            )

            if has_data:
                print(
                    f"SUCCESS: {fund_name} | "
                    f"AUM={aum} | "
                    f"1Y={performance['1Y']} | "
                    f"2Y={performance['2Y']} | "
                    f"3Y={performance['3Y']} | "
                    f"5Y={performance['5Y']}"
                )
                return result

            print("No data extracted.")

        except PlaywrightTimeoutError as exc:
            print(f"Timeout: {exc}")

        except Exception as exc:
            print(f"Error: {exc}")

        if attempt < MAX_RETRIES:
            try:
                page.reload(
                    wait_until="domcontentloaded",
                    timeout=PAGE_TIMEOUT,
                )
            except Exception:
                pass

            time.sleep(1)

    return {
        "Fund": fund_name,
        "AUM (₹ Cr.)": None,
        "1Y (%)": None,
        "2Y (%)": None,
        "3Y (%)": None,
        "5Y (%)": None,
        "Source": fund_url,
    }


# ============================================================
# SCRAPE MULTIPLE FUNDS
# ============================================================

def scrape_funds(selected_funds):
    columns = [
        "Fund",
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
        "Source",
    ]

    if not selected_funds:
        return pd.DataFrame(columns=columns)

    results = []

    with sync_playwright() as p:
        browser = None
        context = None
        page = None

        try:
            print("Starting Chromium...")

            browser = launch_browser(p)
            context = create_context(browser)
            page = context.new_page()

            page.set_default_timeout(ELEMENT_TIMEOUT)
            page.set_default_navigation_timeout(PAGE_TIMEOUT)

            for index, fund_name in enumerate(selected_funds):
                if fund_name not in FUNDS:
                    print(f"Unknown fund: {fund_name}")
                    continue

                fund_url = FUNDS[fund_name]["url"]

                result = scrape_fund(
                    page,
                    fund_name,
                    fund_url,
                )

                results.append(result)

                if index < len(selected_funds) - 1:
                    time.sleep(REQUEST_DELAY)

        finally:
            try:
                if page:
                    page.close()
            except Exception:
                pass

            try:
                if context:
                    context.close()
            except Exception:
                pass

            try:
                if browser:
                    browser.close()
            except Exception:
                pass

    df = pd.DataFrame(results)

    df = df.reindex(columns=columns)

    for column in [
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    return df


# ============================================================
# SAVE EXCEL
# ============================================================

def save_to_excel(df, output_path=OUTPUT_FILE):
    output_path = Path(output_path)

    # Source is intentionally removed from Excel.
    excel_df = df.drop(
        columns=["Source"],
        errors="ignore",
    ).copy()

    required_columns = [
        "Fund",
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]

    excel_df = excel_df.reindex(
        columns=[
            column
            for column in required_columns
            if column in excel_df.columns
        ]
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:
        excel_df.to_excel(
            writer,
            sheet_name="Mutual Funds",
            index=False,
        )

    workbook = load_workbook(output_path)
    worksheet = workbook["Mutual Funds"]

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for row in worksheet.iter_rows(min_row=2):
        if len(row) >= 6:
            row[1].number_format = "#,##0.00"
            row[2].number_format = "0.00"
            row[3].number_format = "0.00"
            row[4].number_format = "0.00"
            row[5].number_format = "0.00"

        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
            )

    widths = {
        "A": 48,
        "B": 18,
        "C": 12,
        "D": 12,
        "E": 12,
        "F": 12,
    }

    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    worksheet.row_dimensions[1].height = 25

    workbook.save(output_path)

    print(f"Excel saved: {output_path}")

    return str(output_path)


# ============================================================
# OPTIONAL COMMAND-LINE TEST
# ============================================================

if __name__ == "__main__":
    print("=" * 70)
    print("MONEYCONTROL MUTUAL FUND SCRAPER")
    print("=" * 70)

    selected_funds = list(FUNDS.keys())

    dataframe = scrape_funds(selected_funds)

    print()
    print(dataframe.to_string(index=False))

    save_to_excel(
        dataframe,
        OUTPUT_FILE,
    )
