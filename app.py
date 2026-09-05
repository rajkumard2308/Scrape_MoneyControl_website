import os

import streamlit as st

from main import (
    FUNDS,
    scrape_funds,
    save_to_excel,
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Mutual Fund Scraper",
    page_icon="📊",
    layout="wide",
)


# ============================================================
# SESSION STATE
# ============================================================

if "results" not in st.session_state:
    st.session_state.results = None

if "result_mode" not in st.session_state:
    st.session_state.result_mode = None


# ============================================================
# TITLE
# ============================================================

st.title(
    "📊 Mutual Fund Performance"
)

st.caption(
    "Moneycontrol performance data"
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "Fund Selection"
    )

    mode = st.radio(
        "Return Type",
        [
            "Lumpsum",
            "SIP",
        ],
        horizontal=True,
    )

    st.divider()

    select_all = st.checkbox(
        "Select all funds"
    )

    fund_names = list(
        FUNDS.keys()
    )

    if select_all:

        selected_funds = fund_names

        st.caption(
            f"{len(selected_funds)} fund(s) selected"
        )

    else:

        selected_funds = st.multiselect(
            "Choose fund(s)",
            fund_names,
        )

        st.caption(
            f"{len(selected_funds)} fund(s) selected"
        )

    st.divider()

    fetch = st.button(
        "🚀 Fetch Fund Data",
        type="primary",
        use_container_width=True,
    )

    clear = st.button(
        "🗑️ Clear Results",
        use_container_width=True,
    )


# ============================================================
# CLEAR
# ============================================================

if clear:

    st.session_state.results = None
    st.session_state.result_mode = None

    st.rerun()


# ============================================================
# FETCH
# ============================================================

if fetch:

    if not selected_funds:

        st.warning(
            "Please select at least one fund."
        )

    else:

        with st.spinner(
            f"Fetching {mode} data for "
            f"{len(selected_funds)} fund(s)..."
        ):

            try:

                df = scrape_funds(
                    selected_funds,
                    mode,
                )

                st.session_state.results = df
                st.session_state.result_mode = mode

            except Exception as e:

                st.error(
                    f"Scraping failed: {e}"
                )

                st.session_state.results = None
                st.session_state.result_mode = None


# ============================================================
# RESULTS
# ============================================================

df = st.session_state.results
result_mode = st.session_state.result_mode


if df is not None:

    # Determine whether anything was actually extracted.
    numeric_columns = [
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]

    has_data = (
        df[numeric_columns]
        .notna()
        .any()
        .any()
    )

    if has_data:

        st.success(
            f"Fetched {len(df)} fund(s)."
        )

    else:

        st.error(
            "Moneycontrol opened, but no fund data "
            "was extracted. Please try again."
        )

    st.subheader(
        f"{result_mode} Performance"
    )

    display_df = df.drop(
        columns=["Source"],
        errors="ignore"
    )

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )

    # ========================================================
    # EXCEL
    # ========================================================

    excel_path = "mutual_funds.xlsx"

    save_to_excel(
        df,
        excel_path,
    )

    with open(
        excel_path,
        "rb"
    ) as file:

        st.download_button(
            "📥 Download Excel",
            data=file,
            file_name="mutual_funds.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

else:

    st.info(
        "Select fund(s), choose Lumpsum or SIP, "
        "and click Fetch Fund Data."
    )