"""Excel ingestion utilities for TenderPro."""

from __future__ import annotations

from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

from modules.mapping import REQUIRED_COLUMNS, detect_columns

SAMPLE_FILES = {
    "Enext_BOQ.xlsx": [
        {"Item No": "1.01", "Item Description": "GI cable tray 300 mm", "UOM": "m", "Qty": 120, "Rate": 42.5, "Supplier": "Enext"},
        {"Item No": "1.02", "Item Description": "LV panel type A", "UOM": "no", "Qty": 2, "Rate": 8200, "Supplier": "Enext"},
        {"Item No": "1.03", "Item Description": "Earthing pit complete", "UOM": "no", "Qty": 8, "Rate": None, "Supplier": "Enext"},
        {"Item No": "1.04", "Item Description": "Copper cable 4C x 16 sq.mm", "UOM": "m", "Qty": 450, "Rate": 18.75, "Supplier": "Enext"},
    ],
    "Valtria_BOQ.xlsx": [
        {"Code": "1.01", "Description": "GI cable tray 300 mm", "Unit": "m", "Quantity": 120, "Unit Rate": 39.9, "Vendor": "Valtria"},
        {"Code": "1.02", "Description": "LV panel type A", "Unit": "no", "Quantity": 2, "Unit Rate": 8750, "Vendor": "Valtria"},
        {"Code": "1.03", "Description": "Earthing pit complete", "Unit": "no", "Quantity": 8, "Unit Rate": 610, "Vendor": "Valtria"},
        {"Code": "1.04", "Description": "Copper cable 4C x 16 sq.mm", "Unit": "m", "Quantity": 450, "Unit Rate": 19.1, "Vendor": "Valtria"},
    ],
}


def create_sample_files(directory: str | Path = "uploads") -> list[Path]:
    """Create sample vendor BOQs when the upload folder has no spreadsheets."""
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
