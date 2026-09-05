import os
import tempfile

import streamlit as st

from main import FUNDS, scrape_funds, save_to_excel


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title="Mutual Fund Scraper",
    page_icon="📊",
    layout="wide",
)


# ============================================================
# SESSION STATE
# ============================================================

if "fund_data" not in st.session_state:
    st.session_state.fund_data = None

if "selected_funds" not in st.session_state:
    st.session_state.selected_funds = []

if "return_type" not in st.session_state:
    st.session_state.return_type = "Lumpsum"


def clear_results():
    st.session_state.fund_data = None
    st.session_state.selected_funds = []


# ============================================================
# SIDEBAR
# ============================================================

names = list(FUNDS.keys())

with st.sidebar:

    st.title("📊 Mutual Fund Scraper")

    st.caption(
        "Moneycontrol performance data"
    )

    st.divider()

    st.subheader(
        "Fund Selection"
    )

    mode = st.radio(
        "Return Type",
        ["Lumpsum", "SIP"],
        horizontal=True,
        key="return_type",
    )

    all_funds = st.checkbox(
        "Select all funds"
    )

    previous = [
        fund
        for fund in st.session_state.selected_funds
        if fund in names
    ]

    selected = st.multiselect(
        "Choose fund(s)",
        names,
        default=(
            names
            if all_funds
            else previous
        ),
    )

    st.session_state.selected_funds = selected

    st.caption(
        f"{len(selected)} fund(s) selected"
    )

    fetch = st.button(
        "🚀 Fetch Fund Data",
        type="primary",
        use_container_width=True,
    )

    st.button(
        "🗑️ Clear Results",
        use_container_width=True,
        on_click=clear_results,
    )

    st.divider()

    st.write(
        "Data collected"
    )

    st.write("• Fund Name")
    st.write("• AUM")
    st.write(
        f"• {mode} Annualised Return — 1Y"
    )
    st.write(
        f"• {mode} Annualised Return — 2Y"
    )
    st.write(
        f"• {mode} Annualised Return — 3Y"
    )
    st.write(
        f"• {mode} Annualised Return — 5Y"
    )


# ============================================================
# MAIN
# ============================================================

st.title(
    "📊 Mutual Fund Performance"
)

st.caption(
    f"{mode} annualised returns with AUM."
)


# ============================================================
# FETCH
# ============================================================

if fetch:

    if not selected:

        st.warning(
            "Please select at least one mutual fund."
        )

    else:

        progress = st.empty()

        try:

            progress.info(
                f"Fetching {len(selected)} fund(s) "
                f"— {mode}..."
            )

            with st.spinner(
                f"Fetching {mode} data from Moneycontrol..."
            ):

                result = scrape_funds(
                    selected,
                    mode
                )

            st.session_state.fund_data = result

            progress.success(
                f"Fetched {len(selected)} fund(s)."
            )

        except Exception as e:

            progress.empty()

            st.error(
                f"Scraping failed: {e}"
            )


# ============================================================
# RESULT
# ============================================================

df = st.session_state.fund_data


if df is None or df.empty:

    st.info(
        "Select fund(s), choose Lumpsum or SIP, "
        "and click Fetch Fund Data."
    )

else:

    st.subheader(
        f"{mode} Performance"
    )

    display = df[
        [
            "Fund",
            "AUM (₹ Cr.)",
            "1Y (%)",
            "2Y (%)",
            "3Y (%)",
            "5Y (%)",
        ]
    ].copy()

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )

    # ========================================================
    # EXCEL
    # ========================================================

    tmp = None

    try:

        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False
        ) as file:

            tmp = file.name

        save_to_excel(
            display,
            tmp
        )

        with open(
            tmp,
            "rb"
        ) as file:

            excel_data = file.read()

        st.download_button(
            "📥 Download Excel",
            data=excel_data,
            file_name=(
                f"mutual_funds_"
                f"{mode.lower()}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    finally:

        if (
            tmp
            and os.path.exists(tmp)
        ):

            try:
                os.remove(tmp)
            except Exception:
                pass