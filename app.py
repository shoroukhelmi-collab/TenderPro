"""TenderPro Streamlit MVP."""

from __future__ import annotations

import streamlit as st

from modules.dashboard import render_dashboard, render_home_kpis
from modules.excel_reader import WorkbookReview, inspect_excel_file, read_reviewed_excels, read_sample_excels
from modules.export import export_comparison
from modules.mapping import REQUIRED_COLUMNS

st.set_page_config(page_title="TenderPro MVP", page_icon="📊", layout="wide")

st.title("TenderPro")
st.caption("Upload supplier BOQ Excel files, review the detected mapping, generate a comparison, and export one workbook.")

if "comparison_generated" not in st.session_state:
    st.session_state.comparison_generated = False

st.markdown(
    """
    **Supported BOQ columns:** item number/code, description, unit/UOM, quantity/qty,
    unit rate/price, and total amount. If total amount is missing, TenderPro calculates
    it from quantity × unit rate.
    """
)

uploaded_files = st.file_uploader(
    "New Comparison / Upload Supplier Files",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Upload one Excel workbook per supplier. Supplier names and column mappings can be edited before comparison.",
)
use_sample_data = st.checkbox("Use sample data", value=False, help="Use two small generic supplier workbooks for a quick demo.")

normalized_data = None
reviews: list[WorkbookReview] = []

if uploaded_files:
    st.subheader("File review")
    st.caption("Confirm each detected supplier name, worksheet, header row, column mapping, and imported row count before generating.")
    for index, uploaded_file in enumerate(uploaded_files, start=1):
        uploaded_file.seek(0)
        review = inspect_excel_file(uploaded_file, f"Supplier {index}")
        with st.expander(f"{review.file_name} — detected as {review.supplier_name}", expanded=True):
            supplier_name = st.text_input("Supplier name", review.supplier_name, key=f"supplier_{index}")
            worksheet = st.selectbox("Worksheet", review.worksheet_names, index=review.worksheet_names.index(review.selected_worksheet), key=f"sheet_{index}")
            header_row = st.number_input("Detected header row", min_value=1, value=review.header_row, step=1, key=f"header_{index}")
            st.write("Detected worksheet names:", ", ".join(review.worksheet_names))
            column_options = [None, *review.columns]
            mapping = {}
            mapping_cols = st.columns(3)
            for field_index, canonical in enumerate(REQUIRED_COLUMNS):
                current = review.column_mapping.get(canonical)
                selected_index = column_options.index(current) if current in column_options else 0
                mapping[canonical] = mapping_cols[field_index % 3].selectbox(
                    canonical.replace("_", " ").title(),
                    column_options,
                    index=selected_index,
                    key=f"mapping_{index}_{canonical}",
                )
            st.metric("Number of imported rows", review.imported_rows)
            reviews.append(
                WorkbookReview(
                    file_name=review.file_name,
                    supplier_name=supplier_name,
                    worksheet_names=review.worksheet_names,
                    selected_worksheet=worksheet,
                    header_row=int(header_row),
                    column_mapping=mapping,
                    imported_rows=review.imported_rows,
                    columns=review.columns,
                )
            )
    if st.button("Generate Comparison", type="primary"):
        st.session_state.comparison_generated = True

    if st.session_state.comparison_generated:
        for uploaded_file in uploaded_files:
            uploaded_file.seek(0)
        normalized_data = read_reviewed_excels(uploaded_files, reviews)
elif use_sample_data:
    normalized_data = read_sample_excels()
    st.session_state.comparison_generated = True
else:
    st.info("Upload supplier Excel files with the button above, or tick **Use sample data** to try the MVP.")

render_home_kpis(normalized_data, len(uploaded_files) if uploaded_files else (2 if use_sample_data else 0))

if normalized_data is not None and st.session_state.comparison_generated:
    if normalized_data.empty:
        st.warning("No BOQ rows were detected. Check that your files contain item numbers or descriptions.")
    else:
        render_dashboard(normalized_data, len(uploaded_files) if uploaded_files else 2)
        st.download_button(
            "Export final comparison to Excel",
            data=export_comparison(normalized_data, len(uploaded_files) if uploaded_files else 2, st.session_state.get("excluded_outliers", set())),
            file_name="tenderpro_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
