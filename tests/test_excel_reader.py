from io import BytesIO

import pandas as pd
from openpyxl import Workbook

from modules.excel_reader import detect_supplier_name, inspect_excel_file, read_excel_file


def _book(rows_by_sheet, name="Acme Tender BOQ Rev 02.xlsx", merges=None):
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in rows_by_sheet.items():
        ws = wb.create_sheet(sheet_name)
        for row in rows:
            ws.append(row)
        for merge in (merges or {}).get(sheet_name, []):
            ws.merge_cells(merge)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    bio.name = name
    return bio


def test_title_rows_above_header_are_excluded_and_header_detected():
    file = _book({"BOQ": [
        ["Project Alpha"],
        ["Main Building Package"],
        ["Item No", "Description", "Unit", "Qty", "Rate", "Amount"],
        ["1.01", "Concrete works", "m3", 10, 20, 200],
    ]})

    review = inspect_excel_file(file)
    data = read_excel_file(file, review.supplier_name, "BOQ", review.worksheets[0].header_row - 1, review.worksheets[0].column_mapping)

    assert review.worksheets[0].header_row == 3
    assert len(data) == 1
    assert data.iloc[0]["description"] == "Concrete works"


def test_section_headings_are_inherited_as_package_not_items():
    file = _book({"BOQ": [
        ["Item", "Description", "Unit", "Quantity", "Unit Rate", "Total"],
        [None, "Section A - Civil Works", None, None, None, None],
        ["A1", "Excavation", "m3", 5, 3, 15],
    ]})

    data = read_excel_file(file)

    assert len(data) == 1
    assert data.iloc[0]["package"] == "Section A - Civil Works"


def test_subtotal_and_total_rows_are_excluded():
    file = _book({"BOQ": [
        ["Code", "Description", "Unit", "Qty", "Rate", "Amount"],
        ["1", "Doors", "ea", 2, 100, 200],
        [None, "Subtotal", None, None, None, 200],
        [None, "Grand Total", None, None, None, 200],
    ]})

    data = read_excel_file(file)

    assert data["description"].tolist() == ["Doors"]


def test_merged_cells_do_not_create_extra_items():
    file = _book({"BOQ": [
        ["Project title", None, None, None, None, None],
        ["Item No", "Description", "UOM", "Qty", "Rate", "Amount"],
        ["1", "Blockwork", "m2", 7, 4, 28],
    ]}, merges={"BOQ": ["A1:F1"]})

    data = read_excel_file(file)

    assert len(data) == 1
    assert data.iloc[0]["item_no"] == "1"


def test_multiple_worksheets_are_reviewed_and_read():
    file = _book({
        "Civil": [["Item", "Description", "Unit", "Qty", "Rate", "Amount"], ["C1", "Concrete", "m3", 1, 2, 2]],
        "MEP": [["Item", "Description", "Unit", "Qty", "Rate", "Amount"], ["M1", "Cable", "m", 3, 4, 12]],
    })

    review = inspect_excel_file(file)
    frames = [read_excel_file(file, review.supplier_name, ws.sheet_name, ws.header_row - 1, ws.column_mapping) for ws in review.worksheets]
    data = pd.concat(frames, ignore_index=True)

    assert {ws.sheet_name for ws in review.worksheets} == {"Civil", "MEP"}
    assert data["description"].tolist() == ["Concrete", "Cable"]


def test_supplier_name_comes_from_file_name_without_revision_noise():
    assert detect_supplier_name("/tmp/Acme Construction BOQ Rev 03 Final.xlsx") == "Acme Construction"


def test_invalid_descriptive_rows_are_excluded():
    file = _book({"BOQ": [
        ["Item No", "Description", "Unit", "Qty", "Rate", "Amount"],
        [None, "Note: rates include delivery", None, None, None, None],
        [None, "General description of works", None, None, None, None],
        ["1", "Valid item", "lot", 1, None, None],
    ]})

    data = read_excel_file(file)

    assert data["description"].tolist() == ["Valid item"]


def test_price_columns_detect_with_currency_suffixes():
    file = _book({"BOQ": [
        ["Ref", "Particulars", "UOM", "BOQ Qty", "Unit Rate (AED)", "Total Amount (AED)"],
        ["1", "Cable containment", "m", 10, 12.5, 125],
    ]})

    review = inspect_excel_file(file)
    mapping = review.worksheets[0].column_mapping
    data = read_excel_file(file, review.supplier_name, "BOQ", review.worksheets[0].header_row - 1, mapping)

    assert mapping["unit_rate"] == "Unit Rate (AED)"
    assert mapping["total_amount"] == "Total Amount (AED)"
    assert data.iloc[0]["unit_rate"] == 12.5
    assert data.iloc[0]["total_amount"] == 125


def test_non_package_headings_are_not_inherited_as_packages():
    file = _book({"BOQ": [
        ["Item", "Description", "Unit", "Quantity", "Unit Rate", "Total"],
        [None, "Method of Measurements", None, None, None, None],
        [None, "General Notes", None, None, None, None],
        [None, "Specifications", None, None, None, None],
        [None, "Scope of Work", None, None, None, None],
        [None, "Section A - Civil Works", None, None, None, None],
        ["A1", "Excavation", "m3", 5, 3, 15],
    ]})

    data = read_excel_file(file)

    assert len(data) == 1
    assert data.iloc[0]["package"] == "Section A - Civil Works"


def test_supplier_name_removes_leading_project_number_revision_boq_and_extension():
    assert detect_supplier_name("3661 CMC04 FP Rev 2.xlsx") == "CMC04 FP"
