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
    st.caption("Confirm detected supplier names and review every worksheet before previewing the imported BOQ rows.")
    for index, uploaded_file in enumerate(uploaded_files, start=1):
        uploaded_file.seek(0)
        review = inspect_excel_file(uploaded_file)
        with st.expander(f"{review.file_name} — detected as {review.supplier_name}", expanded=True):
            supplier_name = st.text_input("Supplier name", review.supplier_name, key=f"supplier_{index}")
            st.write("Detected worksheet names:", ", ".join(review.worksheet_names))
            edited_sheets = []
            for sheet_index, sheet in enumerate(review.worksheets, start=1):
                st.markdown(f"**Worksheet: {sheet.sheet_name}**")
                header_row = st.number_input("Detected header row", min_value=1, value=sheet.header_row, step=1, key=f"header_{index}_{sheet_index}")
                st.write("Detected columns:", ", ".join(sheet.columns) or "None")
                st.write("Package detection:", sheet.package_source)
                if sheet.section_headings:
                    st.write("Detected section headings:", ", ".join(sheet.section_headings))
                column_options = [None, *sheet.columns]
                mapping = {}
                mapping_cols = st.columns(3)
                for field_index, canonical in enumerate(REQUIRED_COLUMNS):
                    current = sheet.column_mapping.get(canonical)
                    selected_index = column_options.index(current) if current in column_options else 0
                    mapping[canonical] = mapping_cols[field_index % 3].selectbox(
                        canonical.replace("_", " ").title(),
                        column_options,
                        index=selected_index,
                        key=f"mapping_{index}_{sheet_index}_{canonical}",
                    )
                st.metric("Imported item count", sheet.imported_rows)
                st.metric("Excluded row count", sheet.excluded_rows)
                edited_sheets.append(type(sheet)(sheet.sheet_name, int(header_row), mapping, sheet.imported_rows, sheet.excluded_rows, sheet.columns, sheet.package_source, sheet.section_headings))
            reviews.append(WorkbookReview(review.file_name, supplier_name, review.worksheet_names, review.selected_worksheet, review.header_row, review.column_mapping, review.imported_rows, review.columns, review.excluded_rows, review.package_source, review.section_headings, edited_sheets))
    for uploaded_file in uploaded_files:
        uploaded_file.seek(0)
    preview_data = read_reviewed_excels(uploaded_files, reviews)
    st.subheader("Imported rows preview")
    st.caption("Review the first 20 commercial BOQ items. The comparison will not be generated until you confirm this preview.")
    st.dataframe(preview_data.head(20), use_container_width=True)
    confirm_preview = st.checkbox("I confirm the preview is ready for comparison", key="confirm_preview")
    if st.button("Generate Comparison", type="primary", disabled=not confirm_preview):
        st.session_state.comparison_generated = True

    if st.session_state.comparison_generated and confirm_preview:
        normalized_data = preview_data
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
