"""Excel ingestion utilities for TenderPro."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

from modules.mapping import REQUIRED_COLUMNS, detect_columns

SAMPLE_SUPPLIERS: dict[str, list[dict[str, object]]] = {
    "Supplier_1.xlsx": [
        {"Item No": "1.01", "Item Description": "Concrete footing", "UOM": "m3", "Qty": 25, "Rate": 95},
        {"Item No": "1.02", "Item Description": "Steel reinforcement", "UOM": "kg", "Qty": 1200, "Rate": 1.8},
        {"Item No": "1.03", "Item Description": "Cable tray", "UOM": "m", "Qty": 80, "Rate": None},
    ],
    "Supplier_2.xlsx": [
        {"Code": "1.01", "Description": "Concrete footing", "Unit": "m3", "Quantity": 25, "Unit Price": 90},
        {"Code": "1.02", "Description": "Steel reinforcement", "Unit": "kg", "Quantity": 1200, "Unit Price": 1.95},
        {"Code": "1.03", "Description": "Cable tray", "Unit": "m", "Quantity": 80, "Unit Price": 22},
    ],
}


@dataclass
class WorkbookReview:
    """Detected workbook metadata shown before comparison generation."""

    file_name: str
    supplier_name: str
    worksheet_names: list[str]
    selected_worksheet: str
    header_row: int
    column_mapping: dict[str, str | None]
    imported_rows: int
    columns: list[str]


def sample_workbooks(directory: str | Path = "uploads") -> list[Path]:
    """Create small generic sample workbooks for local demos."""
    output_dir = Path(directory)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []

    for filename, rows in SAMPLE_SUPPLIERS.items():
        path = output_dir / filename
        frame = pd.DataFrame(rows)
        if "Rate" in frame.columns:
            frame["Amount"] = frame["Qty"] * frame["Rate"]
        if "Unit Price" in frame.columns:
            frame["Total"] = frame["Quantity"] * frame["Unit Price"]
        frame.to_excel(path, index=False)
        paths.append(path)

    return paths


def inspect_excel_file(file: str | Path | BinaryIO, supplier_name: str | None = None) -> WorkbookReview:
    """Inspect workbook sheets, header row, mappings, and importable row count."""
    _rewind(file)
    excel = pd.ExcelFile(file, engine="openpyxl")
    selected_sheet, header_row, columns, mapping, row_count = "", 0, [], {}, 0
    best_score = -1
    for sheet in excel.sheet_names:
        preview = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=25)
        for row_index in range(min(len(preview), 20)):
            candidate_columns = preview.iloc[row_index].fillna("").astype(str).tolist()
            detected = detect_columns(candidate_columns)
            score = len(detected)
            if score > best_score:
                best_score = score
                selected_sheet = sheet
                header_row = row_index
                columns = [str(column) for column in candidate_columns if str(column).strip()]
                mapping = {canonical: detected.get(canonical) for canonical in REQUIRED_COLUMNS}
    if selected_sheet:
        _rewind(file)
        rows = read_excel_file(file, supplier_name or "Supplier", selected_sheet, header_row, mapping)
        row_count = len(rows)
    return WorkbookReview(
        file_name=getattr(file, "name", str(file)),
        supplier_name=supplier_name or Path(getattr(file, "name", "Supplier")).stem.replace("_", " "),
        worksheet_names=excel.sheet_names,
        selected_worksheet=selected_sheet or excel.sheet_names[0],
        header_row=header_row + 1,
        column_mapping=mapping,
        imported_rows=row_count,
        columns=columns,
    )


def read_excel_file(
    file: str | Path | BinaryIO,
    supplier_name: str,
    sheet_name: str | int = 0,
    header_row: int = 0,
    column_mapping: dict[str, str | None] | None = None,
) -> pd.DataFrame:
    """Read one supplier spreadsheet and return canonical BOQ columns."""
    _rewind(file)
    raw = pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name, header=header_row).dropna(how="all")
    detected = column_mapping or detect_columns(raw.columns)

    normalized = pd.DataFrame(index=raw.index)
    for canonical in REQUIRED_COLUMNS:
        source_column = detected.get(canonical)
        normalized[canonical] = raw[source_column] if source_column in raw.columns else pd.NA

    normalized["supplier"] = supplier_name
    for numeric_column in ("quantity", "unit_rate", "total_amount"):
        normalized[numeric_column] = pd.to_numeric(normalized[numeric_column], errors="coerce")

    has_total_inputs = normalized["quantity"].notna() & normalized["unit_rate"].notna()
    missing_total = normalized["total_amount"].isna() & has_total_inputs
    normalized.loc[missing_total, "total_amount"] = normalized.loc[missing_total, "quantity"] * normalized.loc[missing_total, "unit_rate"]

    return normalized.dropna(subset=["item_no", "description"], how="all").reset_index(drop=True)


def _rewind(file: str | Path | BinaryIO) -> None:
    """Rewind file-like uploads when Streamlit reuses their buffers."""
    if hasattr(file, "seek"):
        file.seek(0)


def read_reviewed_excels(files: Iterable[BinaryIO], reviews: list[WorkbookReview]) -> pd.DataFrame:
    """Read Streamlit uploads after the user confirms metadata and mapping."""
    frames = []
    for file, review in zip(files, reviews, strict=False):
        frames.append(read_excel_file(file, review.supplier_name, review.selected_worksheet, review.header_row - 1, review.column_mapping))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=[*REQUIRED_COLUMNS, "supplier"])


def read_uploaded_excels(files: Iterable[BinaryIO]) -> pd.DataFrame:
    """Read Streamlit uploads using detected metadata."""
    reviews = [inspect_excel_file(file, f"Supplier {index}") for index, file in enumerate(files, start=1)]
    return read_reviewed_excels(files, reviews)


def read_sample_excels(directory: str | Path = "uploads") -> pd.DataFrame:
    """Read the bundled generic samples."""
    paths = sample_workbooks(directory)
    frames = [read_excel_file(path, f"Supplier {index}") for index, path in enumerate(paths, start=1)]
    return pd.concat(frames, ignore_index=True)
