import os
import tempfile

import streamlit as st

from main import FUNDS, scrape_funds, save_to_excel


st.set_page_config(
    page_title="Mutual Fund Scraper",
    page_icon="📊",
    layout="wide",
)


# -----------------------------
# Session state
# -----------------------------

if "fund_data" not in st.session_state:
    st.session_state["fund_data"] = None

if "selected_funds" not in st.session_state:
    st.session_state["selected_funds"] = []

if "widget_version" not in st.session_state:
    st.session_state["widget_version"] = 0


def clear_results():
    st.session_state["fund_data"] = None
    st.session_state["selected_funds"] = []
    st.session_state["widget_version"] += 1


# -----------------------------
# Sidebar
# -----------------------------

fund_names = list(FUNDS.keys())
version = st.session_state["widget_version"]

with st.sidebar:
    st.title("📊 Mutual Fund Scraper")
    st.caption("Moneycontrol performance data")
    st.divider()

    st.subheader("Fund Selection")

    select_all = st.checkbox(
        "Select all funds",
        key=f"select_all_{version}",
    )

    if select_all:
        selected_funds = st.multiselect(
            "Choose funds",
            fund_names,
            default=fund_names,
            key=f"fund_selector_all_{version}",
        )
    else:
        previous = [
            fund
            for fund in st.session_state["selected_funds"]
            if fund in fund_names
        ]

        selected_funds = st.multiselect(
            "Choose funds",
            fund_names,
            default=previous,
            key=f"fund_selector_{version}",
        )

    st.caption(f"{len(selected_funds)} funds selected")

    fetch_clicked = st.button(
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

    st.write("Data collected")
    st.write("• AUM")
    st.write("• 1Y Annualised Return")
    st.write("• 2Y Annualised Return")
    st.write("• 3Y Annualised Return")
    st.write("• 5Y Annualised Return")

st.session_state["selected_funds"] = selected_funds


# -----------------------------
# Main page
# -----------------------------

st.title("📊 Mutual Fund Performance")
st.caption(
    "AUM and annualised 1Y, 2Y, 3Y and 5Y returns from Moneycontrol."
)


# -----------------------------
# Fetch
# -----------------------------

if fetch_clicked:
    if not selected_funds:
        st.warning("Please select at least one mutual fund.")
    else:
        try:
            with st.spinner("Fetching data from Moneycontrol..."):
                df = scrape_funds(selected_funds)

            st.session_state["fund_data"] = df

            if df is None or df.empty:
                st.warning("No data was returned.")
            else:
                st.success(f"Completed — {len(df)} funds processed.")

        except Exception as exc:
            st.error(f"Scraping failed: {exc}")


# -----------------------------
# Results
# -----------------------------

df = st.session_state["fund_data"]

if df is None or df.empty:
    st.info(
        "Select funds from the sidebar and click "
        "'Fetch Fund Data'."
    )
else:
    st.subheader("Fund Performance")

    required_columns = [
        "Fund",
        "AUM (₹ Cr.)",
        "1Y (%)",
        "2Y (%)",
        "3Y (%)",
        "5Y (%)",
    ]

    display_df = df.reindex(
        columns=[
            column
            for column in required_columns
            if column in df.columns
        ]
    ).copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Download")

    temp_file = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False,
        ) as tmp:
            temp_file = tmp.name

        # Source is excluded by save_to_excel().
        save_to_excel(display_df, temp_file)

        with open(temp_file, "rb") as file:
            excel_bytes = file.read()

        st.download_button(
            "📥 Download Excel",
            data=excel_bytes,
            file_name="mutual_funds.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    except Exception as exc:
        st.error(f"Could not create Excel file: {exc}")

    finally:
        if temp_file and os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except OSError:
                pass
