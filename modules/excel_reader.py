"""Excel ingestion utilities for TenderPro."""

from __future__ import annotations

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


def read_excel_file(file: str | Path | BinaryIO, supplier_name: str) -> pd.DataFrame:
    """Read one supplier spreadsheet and return canonical BOQ columns."""
    raw = pd.read_excel(file, engine="openpyxl").dropna(how="all")
    detected = detect_columns(raw.columns)

    normalized = pd.DataFrame(index=raw.index)
    for canonical in REQUIRED_COLUMNS:
        source_column = detected.get(canonical)
        normalized[canonical] = raw[source_column] if source_column else pd.NA

    normalized["supplier"] = supplier_name
    for numeric_column in ("quantity", "unit_rate", "total_amount"):
        normalized[numeric_column] = pd.to_numeric(normalized[numeric_column], errors="coerce")

    has_total_inputs = normalized["quantity"].notna() & normalized["unit_rate"].notna()
    missing_total = normalized["total_amount"].isna() & has_total_inputs
    normalized.loc[missing_total, "total_amount"] = normalized.loc[missing_total, "quantity"] * normalized.loc[missing_total, "unit_rate"]

    return normalized.dropna(subset=["item_no", "description"], how="all").reset_index(drop=True)


def read_uploaded_excels(files: Iterable[BinaryIO]) -> pd.DataFrame:
    """Read Streamlit uploads using generic Supplier 1, Supplier 2, ... names."""
    frames = [read_excel_file(file, f"Supplier {index}") for index, file in enumerate(files, start=1)]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=[*REQUIRED_COLUMNS, "supplier"])


def read_sample_excels(directory: str | Path = "uploads") -> pd.DataFrame:
    """Read the bundled generic samples."""
    paths = sample_workbooks(directory)
    frames = [read_excel_file(path, f"Supplier {index}") for index, path in enumerate(paths, start=1)]
    return pd.concat(frames, ignore_index=True)
