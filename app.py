"""TenderPro Streamlit MVP."""

from __future__ import annotations

import streamlit as st

from modules.dashboard import render_dashboard, render_home_kpis
from modules.excel_reader import (
    WorkbookReview,
    inspect_excel_file,
    read_all_excel_sheets,
    read_reviewed_excels,
    read_sample_excels,
)
from modules.export import export_comparison
from modules.mapping import REQUIRED_COLUMNS

st.set_page_config(page_title="TenderPro MVP", page_icon="📊", layout="wide")

st.title("TenderPro")
st.caption("Upload supplier BOQ Excel package files, review the detected mapping, generate a comparison, and export one workbook.")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 0.75rem;
        padding: 0.85rem 1rem;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    div.stButton > button[kind="primary"] {
        width: 100%;
        min-height: 3.25rem;
        font-size: 1.05rem;
        font-weight: 700;
    }
    .supplier-card-title { margin: 0 0 0.15rem; }
    .supplier-card-subtitle { color: #64748b; margin-bottom: 0.5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

if "comparison_generated" not in st.session_state:
    st.session_state.comparison_generated = False
if "supplier_group_count" not in st.session_state:
    st.session_state.supplier_group_count = 1
if "supplier_group_ids" not in st.session_state:
    st.session_state.supplier_group_ids = list(range(1, st.session_state.supplier_group_count + 1))
if "next_supplier_group_id" not in st.session_state:
    st.session_state.next_supplier_group_id = max(st.session_state.supplier_group_ids, default=0) + 1

st.markdown(
    """
    **Supported BOQ columns:** item number/code, description, unit/UOM, quantity/qty,
    unit rate/price, and total amount. If total amount is missing, TenderPro calculates
    it from quantity × unit rate.
    """
)


def _add_supplier_group() -> None:
    st.session_state.supplier_group_ids.append(st.session_state.next_supplier_group_id)
    st.session_state.next_supplier_group_id += 1
    st.session_state.supplier_group_count = len(st.session_state.supplier_group_ids)
    st.session_state.comparison_generated = False


def _remove_supplier_group(group_id: int) -> None:
    if len(st.session_state.supplier_group_ids) <= 1:
        return
    st.session_state.supplier_group_ids = [
        supplier_id for supplier_id in st.session_state.supplier_group_ids if supplier_id != group_id
    ]
    st.session_state.supplier_group_count = len(st.session_state.supplier_group_ids)
    st.session_state.comparison_generated = False


def _package_detection_summary(files: list, supplier_name: str) -> tuple[list[str], list[dict[str, object]], int, bool]:
    package_names_by_file: list[tuple[str, str | None]] = []
    undetected_files: list[str] = []
    warnings: list[dict[str, object]] = []
    imported_item_count = 0
    commercial_columns_found = False

    for file in files:
        try:
            file.seek(0)
            detected_rows = read_all_excel_sheets(file, supplier_name)
            imported_item_count += len(detected_rows)
            commercial_columns_found = commercial_columns_found or (
                not detected_rows.empty
                and (detected_rows["unit_rate"].notna().any() or detected_rows["total_amount"].notna().any())
            )
            packages = (
                sorted(
                    detected_rows["package"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .replace("", None)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "package" in detected_rows
                else []
            )
        except Exception as exc:  # Surface upload-card detection issues without changing import review behavior.
            packages = []
            warnings.append({"type": "error", "message": f"Could not inspect {file.name}: {exc}"})
        finally:
            file.seek(0)

        if packages:
            for package in packages:
                package_names_by_file.append((file.name, package))
        else:
            package_names_by_file.append((file.name, None))
            undetected_files.append(file.name)

    detected_package_names = [package for _, package in package_names_by_file if package]
    duplicates = sorted({package for package in detected_package_names if detected_package_names.count(package) > 1})
    for package in duplicates:
        matching_files = [file_name for file_name, detected in package_names_by_file if detected == package]
        warnings.append({"type": "duplicate", "package": package, "files": matching_files})
    for file_name in undetected_files:
        warnings.append({"type": "missing_package", "file": file_name})

    return sorted(set(detected_package_names)), warnings, imported_item_count, commercial_columns_found


st.subheader("New Comparison / Upload Supplier Package Files")
st.caption("Create one group per supplier, enter the supplier name, then upload all package files belonging to that supplier.")
home_kpi_placeholder = st.container()
supplier_groups = []
for group_position, group_id in enumerate(st.session_state.supplier_group_ids, start=1):
    with st.container(border=True):
        existing_name = st.session_state.get(f"supplier_group_name_{group_id}", "")
        supplier_display_name = (existing_name or f"Supplier {group_position}").strip()
        title_col, action_col = st.columns([0.82, 0.18], vertical_alignment="center")
        title_col.markdown(f"### {supplier_display_name}")
        title_col.caption(f"Supplier group {group_position}")
        remove_disabled = st.session_state.supplier_group_count <= 1
        action_col.button(
            "Remove",
            key=f"remove_supplier_group_{group_id}",
            disabled=remove_disabled,
            help="At least one supplier group is required." if remove_disabled else "Remove this supplier group.",
            on_click=_remove_supplier_group,
            args=(group_id,),
        )

        name_col, upload_col = st.columns([0.35, 0.65], vertical_alignment="top")
        supplier_name = name_col.text_input(
            "Supplier Name",
            key=f"supplier_group_name_{group_id}",
            placeholder=f"Supplier {group_position}",
        )
        supplier_display_name = (supplier_name or f"Supplier {group_position}").strip()
        files = upload_col.file_uploader(
            "Upload Excel package files",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"supplier_group_files_{group_id}",
            help="Upload all package workbooks for this supplier. The supplier name is taken only from this group field.",
        )
        files = files or []

        detected_packages: list[str] = []
        package_warnings: list[dict[str, object]] = []
        imported_item_count = 0
        commercial_columns_found = False
        if files:
            detected_packages, package_warnings, imported_item_count, commercial_columns_found = _package_detection_summary(files, supplier_display_name)

        files_col, package_col, status_col = st.columns([0.32, 0.38, 0.30], vertical_alignment="top")
        with files_col.container(border=True):
            st.markdown("**📁 Uploaded Files**")
            st.metric("", f"{len(files)} File{'s' if len(files) != 1 else ''}", label_visibility="collapsed")
            if files:
                for file in files:
                    st.write(f"✓ {file.name}")
            else:
                st.info("No files uploaded yet.")

        with package_col.container(border=True):
            st.markdown("**Detected Packages**")
            if detected_packages:
                for package in detected_packages:
                    st.write(f"✓ {package}")
            elif files:
                st.info("No package names were detected from the uploaded files yet.")
            else:
                st.info("Upload files to detect package names.")

        with status_col.container(border=True):
            st.markdown("**Supplier Status**")
            status_rows = [
                (bool(files), "Files Uploaded", "No Files Uploaded"),
                (bool(detected_packages), "Packages Detected", "Package Detection Incomplete"),
                (commercial_columns_found, "Commercial Columns Found", "Missing Commercial Columns"),
                (bool(files) and bool(detected_packages) and commercial_columns_found, "Ready for Comparison", "Not Ready for Comparison"),
            ]
            for ok, success, failure in status_rows:
                icon = "✓" if ok else ("⚠" if "Package" in failure else "❌")
                st.write(f"{icon} {success if ok else failure}")

        for warning in package_warnings:
            if warning["type"] == "duplicate":
                with st.container(border=True):
                    st.warning("⚠ Duplicate Package")
                    st.markdown(f"**{warning['package']}**")
                    st.write("Found in:")
                    for file_name in warning["files"]:
                        st.write(f"• {file_name}")
            elif warning["type"] == "missing_package":
                st.warning(f"⚠ Package detection incomplete for {warning['file']}.")
            else:
                st.warning(str(warning["message"]))

        with st.container(border=True):
            st.markdown("**Supplier Summary**")
            summary_cols = st.columns(5)
            summary_cols[0].metric("Supplier", supplier_display_name)
            summary_cols[1].metric("Uploaded Files", len(files))
            summary_cols[2].metric("Detected Packages", len(detected_packages))
            summary_cols[3].metric("Imported Commercial Items", imported_item_count)
            summary_cols[4].metric("Ready", "YES" if (files and detected_packages and commercial_columns_found) else "NO")

        supplier_groups.append({
            "name": supplier_display_name,
            "files": files,
            "detected_packages": detected_packages,
            "imported_items": imported_item_count,
            "ready": bool(files) and bool(detected_packages) and commercial_columns_found,
        })

with home_kpi_placeholder:
    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Total Suppliers", len(supplier_groups))
    kpi_cols[1].metric("Total Uploaded Files", sum(len(group["files"]) for group in supplier_groups))
    kpi_cols[2].metric("Detected Packages", sum(len(group["detected_packages"]) for group in supplier_groups))
    kpi_cols[3].metric("Imported Commercial Items", sum(group["imported_items"] for group in supplier_groups))
    kpi_cols[4].metric("Ready Suppliers", sum(1 for group in supplier_groups if group["ready"]))

st.button("+ Add Supplier", on_click=_add_supplier_group, type="secondary")
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
    st.dataframe(preview_data.head(20), use_container_width=True, hide_index=True)
    confirm_preview = st.checkbox("I confirm the preview is ready for comparison", key="confirm_preview")
    missing_all_prices = preview_data.empty or (preview_data["unit_rate"].isna().all() and preview_data["total_amount"].isna().all())
    if missing_all_prices:
        st.error("Comparison cannot be generated because no Unit Rate or Total Amount columns are mapped in the included files.")
    invalid_suppliers = [group["name"] for group in supplier_groups if not group["ready"]]
    generate_disabled = not confirm_preview or missing_all_prices or bool(invalid_suppliers)
    if missing_all_prices:
        generate_help = "Map at least one Unit Rate or Total Amount column before generating the comparison."
    elif invalid_suppliers:
        generate_help = "Complete missing uploads, package detection, or commercial columns for: " + ", ".join(invalid_suppliers)
    elif not confirm_preview:
        generate_help = "Confirm the imported rows preview before generating the comparison."
    else:
        generate_help = "Generate the supplier comparison."
    if st.button("Generate Comparison", type="primary", disabled=generate_disabled, help=generate_help):
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
