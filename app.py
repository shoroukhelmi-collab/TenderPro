"""TenderPro Streamlit MVP."""

from __future__ import annotations

import streamlit as st

from modules.dashboard import render_dashboard
from modules.excel_reader import read_sample_excels, read_uploaded_excels
from modules.export import export_comparison

st.set_page_config(page_title="TenderPro MVP", page_icon="📊", layout="wide")

st.title("TenderPro")
st.caption("Upload supplier BOQ Excel files, compare line items, identify lowest prices and missing prices, then export the result.")

with st.sidebar:
    st.header("Supplier files")
    uploaded_files = st.file_uploader(
        "Upload multiple Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="TenderPro automatically detects common BOQ columns. Supplier names are shown generically as Supplier 1, Supplier 2, and so on.",
    )
    use_sample_data = st.checkbox("Use sample data", value=False, help="Use two small generic supplier workbooks for a quick demo.")

st.markdown(
    """
    **Supported BOQ columns:** item number/code, description, unit/UOM, quantity/qty,
    unit rate/price, and total amount. If total amount is missing, TenderPro calculates
    it from quantity × unit rate.
    """
)

if uploaded_files:
    normalized_data = read_uploaded_excels(uploaded_files)
elif use_sample_data:
    normalized_data = read_sample_excels()
else:
    normalized_data = None
    st.info("Upload at least one supplier Excel file, or tick **Use sample data** to try the MVP.")

if normalized_data is not None:
    if normalized_data.empty:
        st.warning("No BOQ rows were detected. Check that your files contain item numbers or descriptions.")
    else:
        render_dashboard(normalized_data)
        st.download_button(
            "Export final comparison to Excel",
            data=export_comparison(normalized_data),
            file_name="tenderpro_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
