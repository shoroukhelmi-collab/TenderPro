"""TenderPro Streamlit MVP."""

from __future__ import annotations

import streamlit as st

from modules.dashboard import render_dashboard, render_home_kpis
from modules.excel_reader import WorkbookReview, inspect_excel_file, read_reviewed_excels, read_sample_excels
from modules.export import export_comparison
from modules.mapping import REQUIRED_COLUMNS

st.set_page_config(page_title="TenderPro MVP", page_icon="📊", layout="wide")

st.title("TenderPro")
st.caption("Upload supplier BOQ Excel package files, review the detected mapping, generate a comparison, and export one workbook.")

if "comparison_generated" not in st.session_state:
    st.session_state.comparison_generated = False
if "supplier_group_count" not in st.session_state:
    st.session_state.supplier_group_count = 1

st.markdown(
    """
    **Supported BOQ columns:** item number/code, description, unit/UOM, quantity/qty,
    unit rate/price, and total amount. If total amount is missing, TenderPro calculates
    it from quantity × unit rate.
    """
)


def _add_supplier_group() -> None:
    st.session_state.supplier_group_count += 1
    st.session_state.comparison_generated = False

st.subheader("New Comparison / Upload Supplier Package Files")
st.caption("Create one group per supplier, enter the supplier name, then upload all package files belonging to that supplier.")
supplier_groups = []
for group_index in range(1, st.session_state.supplier_group_count + 1):
    with st.container(border=True):
        st.markdown(f"**Supplier {group_index}**")
        supplier_name = st.text_input("Supplier Name", key=f"supplier_group_name_{group_index}", placeholder=f"Supplier {group_index}")
        files = st.file_uploader(
            "Upload multiple Excel files",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"supplier_group_files_{group_index}",
            help="Upload all package workbooks for this supplier. The supplier name is taken only from this group field.",
        )
        if files:
            st.write("Files:", ", ".join(file.name for file in files))
        supplier_groups.append({"name": (supplier_name or f"Supplier {group_index}").strip(), "files": files or []})

st.button("Add another supplier", on_click=_add_supplier_group)
use_sample_data = st.checkbox("Use sample data", value=False, help="Use two small generic supplier workbooks for a quick demo.")

normalized_data = None
reviews: list[WorkbookReview] = []
grouped_uploads = [(group["name"], file) for group in supplier_groups for file in group["files"]]
uploaded_files = [file for _, file in grouped_uploads]

if grouped_uploads:
    st.subheader("File review")
    st.caption("Confirm each supplier group and review every worksheet before previewing the imported BOQ rows.")
    summary_rows = []
    for index, (supplier_name, uploaded_file) in enumerate(grouped_uploads, start=1):
        uploaded_file.seek(0)
        review = inspect_excel_file(uploaded_file, supplier_name=supplier_name)
        with st.expander(f"{supplier_name} — {review.file_name}", expanded=True):
            st.write("Supplier name:", supplier_name)
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
                package_current = sheet.column_mapping.get("package")
                package_index = column_options.index(package_current) if package_current in column_options else 0
                mapping["package"] = st.selectbox(
                    "Package Mapping",
                    column_options,
                    index=package_index,
                    key=f"mapping_{index}_{sheet_index}_package",
                    help="Choose a package/section column, or leave blank to inherit valid section headings.",
                )
                if not any(mapping.get(col) for col in ("unit_rate", "total_amount")):
                    st.warning("No Unit Rate or Total Amount column detected for this worksheet. It will be excluded from comparison until one is mapped.")
                st.metric("Imported item count", sheet.imported_rows)
                st.metric("Excluded row count", sheet.excluded_rows)
                package_source = mapping["package"] or "Section headings"
                summary_rows.append({
                    "Supplier name": supplier_name,
                    "File name": review.file_name,
                    "Worksheet": sheet.sheet_name,
                    "Header row": int(header_row),
                    "Imported rows": sheet.imported_rows,
                    "Excluded rows": sheet.excluded_rows,
                    "Detected Unit Rate column": mapping.get("unit_rate") or "Missing",
                    "Detected Total Amount column": mapping.get("total_amount") or "Missing",
                    "Detected Package source": package_source,
                })
                edited_sheets.append(type(sheet)(sheet.sheet_name, int(header_row), mapping, sheet.imported_rows, sheet.excluded_rows, sheet.columns, package_source, sheet.section_headings))
            reviews.append(WorkbookReview(review.file_name, supplier_name, review.worksheet_names, review.selected_worksheet, review.header_row, review.column_mapping, review.imported_rows, review.columns, review.excluded_rows, review.package_source, review.section_headings, edited_sheets))
    st.subheader("Import detection summary")
    st.dataframe(summary_rows, use_container_width=True)
    if any(row["Detected Unit Rate column"] == "Missing" and row["Detected Total Amount column"] == "Missing" for row in summary_rows):
        st.warning("One or more worksheets have no detected price columns and will not be included in the comparison until Unit Rate or Total Amount is mapped.")
    for uploaded_file in uploaded_files:
        uploaded_file.seek(0)
    preview_data = read_reviewed_excels(uploaded_files, reviews)
    st.subheader("Imported rows preview")
    st.caption("Review the first 20 commercial BOQ items. The comparison will not be generated until you confirm this preview.")
    st.dataframe(preview_data.head(20), use_container_width=True)
    confirm_preview = st.checkbox("I confirm the preview is ready for comparison", key="confirm_preview")
    missing_all_prices = preview_data.empty or (preview_data["unit_rate"].isna().all() and preview_data["total_amount"].isna().all())
    if missing_all_prices:
        st.error("Comparison cannot be generated because no Unit Rate or Total Amount columns are mapped in the included files.")
    if st.button("Generate Comparison", type="primary", disabled=(not confirm_preview or missing_all_prices)):
        st.session_state.comparison_generated = True

    if st.session_state.comparison_generated and confirm_preview:
        normalized_data = preview_data
elif use_sample_data:
    normalized_data = read_sample_excels()
    st.session_state.comparison_generated = True
else:
    st.info("Create supplier groups and upload Excel package files with the controls above, or tick **Use sample data** to try the MVP.")

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
