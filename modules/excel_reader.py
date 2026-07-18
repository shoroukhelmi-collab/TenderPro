"""Excel ingestion utilities for TenderPro."""

from __future__ import annotations

from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

from modules.mapping import REQUIRED_COLUMNS, detect_columns

SAMPLE_FILES = {
    "Supplier_A_BOQ.xlsx": [
        {"Item No": "1.01", "Item Description": "Sample item alpha", "UOM": "unit", "Qty": 120, "Rate": 42.5, "Supplier": "Supplier A"},
        {"Item No": "1.02", "Item Description": "Sample item beta", "UOM": "set", "Qty": 2, "Rate": 8200, "Supplier": "Supplier A"},
        {"Item No": "1.03", "Item Description": "Sample item gamma", "UOM": "unit", "Qty": 8, "Rate": None, "Supplier": "Supplier A"},
        {"Item No": "1.04", "Item Description": "Sample item delta", "UOM": "m", "Qty": 450, "Rate": 18.75, "Supplier": "Supplier A"},
    ],
    "Supplier_B_BOQ.xlsx": [
        {"Code": "1.01", "Description": "Sample item alpha", "Unit": "unit", "Quantity": 120, "Unit Rate": 39.9, "Vendor": "Supplier B"},
        {"Code": "1.02", "Description": "Sample item beta", "Unit": "set", "Quantity": 2, "Unit Rate": 8750, "Vendor": "Supplier B"},
        {"Code": "1.03", "Description": "Sample item gamma", "Unit": "unit", "Quantity": 8, "Unit Rate": 610, "Vendor": "Supplier B"},
        {"Code": "1.04", "Description": "Sample item delta", "Unit": "m", "Quantity": 450, "Unit Rate": 19.1, "Vendor": "Supplier B"},
    ],
}


def create_sample_files(directory: str | Path = "uploads") -> list[Path]:
    """Create sample supplier BOQs when the upload folder has no spreadsheets."""
    upload_dir = Path(directory)
    upload_dir.mkdir(parents=True, exist_ok=True)
    existing = list(upload_dir.glob("*.xlsx"))
    if existing:
        return existing

    created: list[Path] = []
    for filename, rows in SAMPLE_FILES.items():
        frame = pd.DataFrame(rows)
        if "Rate" in frame.columns:
            frame["Amount"] = frame["Qty"] * frame["Rate"]
        if "Unit Rate" in frame.columns:
            frame["Total"] = frame["Quantity"] * frame["Unit Rate"]
        output_path = upload_dir / filename
        frame.to_excel(output_path, index=False)
        created.append(output_path)
    return created


def save_uploaded_files(files: Iterable[BinaryIO], directory: str | Path = "uploads") -> list[Path]:
    """Persist Streamlit uploaded files so they can be read by OpenPyXL/Pandas."""
    upload_dir = Path(directory)
    upload_dir.mkdir(parents=True, exist_ok=True)
    saved_paths: list[Path] = []
    for file in files:
        target = upload_dir / Path(file.name).name
        target.write_bytes(file.getbuffer())
        saved_paths.append(target)
    return saved_paths


def read_excel_file(path: str | Path) -> pd.DataFrame:
    """Read one BOQ spreadsheet and return canonical TenderPro columns."""
    source_path = Path(path)
    raw = pd.read_excel(source_path, engine="openpyxl")
    raw = raw.dropna(how="all")
    detected = detect_columns(raw.columns)

    normalized = pd.DataFrame()
    for canonical in REQUIRED_COLUMNS:
        source_column = detected.get(canonical)
        normalized[canonical] = raw[source_column] if source_column else pd.NA

    fallback_vendor = source_path.stem.replace("_", " ").replace(" BOQ", "").strip()
    normalized["vendor"] = normalized["vendor"].fillna(fallback_vendor)
    normalized["source_file"] = source_path.name

    for numeric_column in ("quantity", "unit_rate", "total_amount"):
        normalized[numeric_column] = pd.to_numeric(normalized[numeric_column], errors="coerce")

    missing_total = normalized["total_amount"].isna() & normalized["quantity"].notna() & normalized["unit_rate"].notna()
    normalized.loc[missing_total, "total_amount"] = normalized.loc[missing_total, "quantity"] * normalized.loc[missing_total, "unit_rate"]

    return normalized.dropna(subset=["item_no", "description"], how="all")


def read_all_excels(directory: str | Path = "uploads") -> pd.DataFrame:
    """Read all Excel files from a directory into one master dataframe."""
    paths = sorted(Path(directory).glob("*.xlsx"))
    if not paths:
        paths = create_sample_files(directory)
    frames = [read_excel_file(path) for path in paths]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=REQUIRED_COLUMNS)
