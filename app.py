import time
from datetime import datetime

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

st.markdown(
    """
    <style>
        .block-container {padding-top: 2rem;}
        div[data-testid="stMetricValue"] {font-size: 1.6rem;}

        /* Compact sidebar: less top padding, tighter gaps between
           widgets, and slimmer dividers. */
        section[data-testid="stSidebar"] .block-container {
            padding-top: 1.5rem;
        }
        section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
            gap: 0.5rem;
        }
        section[data-testid="stSidebar"] hr {
            margin: 0.4rem 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "results": None,
    "result_mode": None,
    "fetch_seconds": None,
    "fetched_at": None,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# HEADER
# ============================================================

title_col, badge_col = st.columns([5, 1])

with title_col:
    st.title("📊 Mutual Fund Performance")
    st.caption(
        "Live Lumpsum & SIP annualised returns and AUM, "
        "sourced from Moneycontrol."
    )

with badge_col:
    if st.session_state.fetched_at:
        st.metric(
            "Last fetched",
            st.session_state.fetched_at,
        )

st.divider()


# ============================================================
# SIDEBAR — CONTROLS
# ============================================================

with st.sidebar:

    st.header("Fund Selection")

    mode = st.radio(
        "Return type",
        ["Lumpsum", "SIP"],
        horizontal=True,
    )

    fund_names = list(FUNDS.keys())

    select_all = st.checkbox("Select all funds")

    if select_all:
        selected_funds = fund_names
    else:
        selected_funds = st.multiselect(
            "Choose funds",
            fund_names,
            placeholder="Search and select funds...",
        )

    st.caption(f"**{len(selected_funds)}** of {len(fund_names)} funds selected")

    fetch = st.button(
        "🚀 Fetch Fund Data",
        type="primary",
        use_container_width=True,
        disabled=len(selected_funds) == 0,
    )

    clear = st.button(
        "🗑️ Clear Results",
        use_container_width=True,
        disabled=st.session_state.results is None,
    )

    st.divider()

    st.caption(
        "Data is scraped on demand and reflects Moneycontrol's most "
        "recently published figures at fetch time — not real-time NAV."
    )


# ============================================================
# CLEAR
# ============================================================

if clear:
    st.session_state.results = None
    st.session_state.result_mode = None
    st.session_state.fetch_seconds = None
    st.session_state.fetched_at = None
    st.rerun()


# ============================================================
# FETCH
# ============================================================

if fetch:

    if not selected_funds:
        st.warning("Please select at least one fund.")

    else:

        with st.spinner(
            f"Fetching {mode} data for {len(selected_funds)} funds... "
            "this can take a little while for larger selections."
        ):

            try:
                started = time.perf_counter()

                df = scrape_funds(
                    selected_funds,
                    mode,
                )

                elapsed = time.perf_counter() - started

                st.session_state.results = df
                st.session_state.result_mode = mode
                st.session_state.fetch_seconds = elapsed
                st.session_state.fetched_at = datetime.now().strftime(
                    "%d %b, %H:%M"
                )

            except Exception as e:

                st.error(f"Scraping failed: {e}")

                st.session_state.results = None
                st.session_state.result_mode = None
                st.session_state.fetch_seconds = None
                st.session_state.fetched_at = None


# ============================================================
# RESULTS
# ============================================================

df = st.session_state.results
result_mode = st.session_state.result_mode

if df is not None:

    # Numeric columns are derived from whatever the DataFrame actually
    # has, since "Both" mode returns a different set of columns
    # (Lumpsum + SIP side by side) than "Lumpsum" or "SIP" alone.
    numeric_columns = [
        column for column in df.columns
        if column not in ("Fund", "Source")
    ]

    has_data = (
        df[numeric_columns]
        .notna()
        .any()
        .any()
    )

    funds_with_data = int(
        df[numeric_columns].notna().any(axis=1).sum()
    )

    # ------------------------------------------------------
    # Summary row
    # ------------------------------------------------------

    metric_cols = st.columns(4)

    metric_cols[0].metric("Funds requested", len(df))
    metric_cols[1].metric("Funds with data", funds_with_data)
    metric_cols[2].metric(
        "Success rate",
        f"{(funds_with_data / len(df) * 100) if len(df) else 0:.0f}%",
    )

    if st.session_state.fetch_seconds is not None:
        metric_cols[3].metric(
            "Fetch time",
            f"{st.session_state.fetch_seconds:.1f}s",
        )

    if has_data:
        if funds_with_data < len(df):
            st.warning(
                f"{len(df) - funds_with_data} funds returned no data. "
                "Moneycontrol may be slow to respond for those — try "
                "fetching them again."
            )
        else:
            st.success(f"Successfully fetched {len(df)} funds.")
    else:
        st.error(
            "Moneycontrol opened, but no fund data was extracted. "
            "Please try again."
        )

    st.divider()

    st.subheader(f"{result_mode} Performance")

    display_df = df.drop(columns=["Source"], errors="ignore")

    # Format numbers nicely: AUM as ₹ Cr, everything else as %.
    column_config = {}

    if "AUM (₹ Cr.)" in display_df.columns:
        column_config["AUM (₹ Cr.)"] = st.column_config.NumberColumn(
            "AUM (₹ Cr.)",
            format="₹%,.0f",
        )

    for column in display_df.columns:
        if column.endswith("(%)"):
            column_config[column] = st.column_config.NumberColumn(
                column,
                format="%.2f",
            )

    if "Fund" in display_df.columns:
        column_config["Fund"] = st.column_config.TextColumn(
            "Fund",
            width="large",
        )

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
    )

    # ========================================================
    # DOWNLOADS
    # ========================================================

    excel_path = "mutual_funds.xlsx"
    save_to_excel(df, excel_path)

    download_cols = st.columns(2)

    with open(excel_path, "rb") as file:
        download_cols[0].download_button(
            "📥 Download Excel",
            data=file,
            file_name="mutual_funds.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    download_cols[1].download_button(
        "📄 Download CSV",
        data=display_df.to_csv(index=False).encode("utf-8"),
        file_name="mutual_funds.csv",
        mime="text/csv",
        use_container_width=True,
    )

else:

    st.info(
        "Select funds in the sidebar, choose Lumpsum or SIP, "
        "then click **Fetch Fund Data**."
    )