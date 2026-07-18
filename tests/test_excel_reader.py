"""Unit tests for the TenderPro Excel reader."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from modules.excel_reader import ExcelReader


def _workbook_with_sheets(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Priced BOQ"
    sheet.append([None, None, None, None])
    sheet.append(["Tender", "Tender", "Tender", "Tender"])
    sheet.merge_cells("A2:D2")
    sheet.append(["Item No", "Description", "Qty", "Rate"])
    sheet.append(["1.01", "Cable tray", 2, 10])
    sheet.append(["1.02", "Panel", 1, 25])

    second = workbook.create_sheet("Alternative")
    second.append(["Supplier", "Code", "Scope", "Quantity", "Unit Price"])
    second.append(["Valtria", "2.01", "Earthing", 3, 7])

    workbook.create_sheet("Empty")
    workbook.save(path)


def test_load_reads_every_non_empty_worksheet_and_preserves_order(tmp_path: Path) -> None:
    """The reader combines all worksheets in source row order."""
    workbook_path = tmp_path / "vendor.xlsx"
    _workbook_with_sheets(workbook_path)

    frame = ExcelReader().load(workbook_path)

    assert list(frame["item_no"]) == ["1.01", "1.02", "2.01"]
    assert list(frame["description"]) == ["Cable tray", "Panel", "Earthing"]
    assert list(frame["total_amount"]) == [20, 25, 21]
    assert list(frame["worksheet"]) == ["Priced BOQ", "Priced BOQ", "Alternative"]


def test_load_accepts_multiple_files(tmp_path: Path) -> None:
    """The reader accepts an iterable of workbook paths."""
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    pd.DataFrame([{"Item No": "A", "Description": "First"}]).to_excel(first, index=False)
    pd.DataFrame([{"Item No": "B", "Description": "Second"}]).to_excel(second, index=False)

    frame = ExcelReader().load([first, second])

    assert list(frame["item_no"]) == ["A", "B"]
    assert list(frame["source_file"]) == ["first.xlsx", "second.xlsx"]


def test_load_logs_errors_without_crashing(tmp_path: Path, caplog) -> None:  # type: ignore[no-untyped-def]
    """Invalid files are logged and produce an empty dataframe."""
    invalid = tmp_path / "invalid.xlsx"
    invalid.write_text("not an excel file")

    with caplog.at_level(logging.ERROR):
        frame = ExcelReader().load(invalid)

    assert frame.empty
    assert "Failed to open Excel file" in caplog.text


def test_missing_header_row_is_logged(tmp_path: Path, caplog) -> None:  # type: ignore[no-untyped-def]
    """Sheets without recognizable headers are skipped with a log entry."""
    path = tmp_path / "no_headers.xlsx"
    pd.DataFrame([["notes", "only"], ["no", "boq"]]).to_excel(path, index=False, header=False)

    with caplog.at_level(logging.ERROR):
        frame = ExcelReader().load(path)

    assert frame.empty
    assert "Could not detect a header row" in caplog.text
