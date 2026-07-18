"""Excel ingestion utilities for TenderPro."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

from modules.mapping import REQUIRED_COLUMNS, detect_columns

LOGGER = logging.getLogger(__name__)

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


class ExcelReader:
    """Read TenderPro BOQ Excel workbooks into one normalized dataframe.

    The reader accepts one or more `.xlsx` files, reads every worksheet in each
    workbook, auto-detects the header row, forward-fills merged-cell values, and
    logs invalid or unreadable sheets without raising to the caller.
    """

    def __init__(self, logger: logging.Logger | None = None) -> None:
        """Initialize the reader.

        Args:
            logger: Optional logger used for recoverable import errors.
        """
        self.logger = logger or LOGGER

    def load(self, files: str | Path | Iterable[str | Path]) -> pd.DataFrame:
        """Load one or multiple Excel files and return normalized TenderPro rows.

        Args:
            files: A single workbook path or an iterable of workbook paths.

        Returns:
            A dataframe with TenderPro's canonical columns plus `source_file`
            and `worksheet`. Unreadable files and empty sheets are skipped.
        """
        paths = self._coerce_paths(files)
        frames: list[pd.DataFrame] = []

        for path in paths:
            try:
                workbook = pd.ExcelFile(path, engine="openpyxl")
            except Exception as exc:  # pragma: no cover - exact parser errors vary
                self.logger.exception("Failed to open Excel file %s: %s", path, exc)
                continue

            for sheet_name in workbook.sheet_names:
                sheet_frame = self._read_sheet(workbook, sheet_name, path)
                if not sheet_frame.empty:
                    frames.append(sheet_frame)

        columns = [*REQUIRED_COLUMNS, "source_file", "worksheet"]
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=columns)

    def _read_sheet(self, workbook: pd.ExcelFile, sheet_name: str, path: Path) -> pd.DataFrame:
        """Read and normalize one worksheet, returning an empty frame on errors."""
        try:
            raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
        except Exception as exc:  # pragma: no cover - exact parser errors vary
            self.logger.exception("Failed to read sheet %s in %s: %s", sheet_name, path, exc)
            return pd.DataFrame()

        raw = raw.dropna(how="all")
        if raw.empty:
            self.logger.info("Skipping empty sheet %s in %s", sheet_name, path)
            return pd.DataFrame()

        raw = raw.ffill(axis=0)
        header_index = self._detect_header_row(raw)
        if header_index is None:
            self.logger.error("Could not detect a header row in sheet %s of %s", sheet_name, path)
            return pd.DataFrame()

        headers = raw.loc[header_index].tolist()
        data = raw.loc[header_index + 1 :].copy()
        data.columns = [self._stringify_header(header, index) for index, header in enumerate(headers)]
        data = data.dropna(how="all")
        if data.empty:
            self.logger.info("Sheet %s in %s has headers but no data", sheet_name, path)
            return pd.DataFrame()

        normalized = self._normalize(data, path, sheet_name)
        return normalized.dropna(subset=["item_no", "description"], how="all")

    def _detect_header_row(self, frame: pd.DataFrame) -> int | None:
        """Return the row index that best matches TenderPro's known headers."""
        best_index: int | None = None
        best_score = 0
        for index, row in frame.iterrows():
            columns = [self._stringify_header(value, position) for position, value in enumerate(row.tolist())]
            score = len(detect_columns(columns))
            if score > best_score:
                best_index = int(index)
                best_score = score
        return best_index if best_score >= 2 else None

    def _normalize(self, raw: pd.DataFrame, source_path: Path, sheet_name: str) -> pd.DataFrame:
        """Normalize vendor-specific sheet columns to TenderPro's canonical schema."""
        detected = detect_columns(raw.columns)
        normalized = pd.DataFrame(index=raw.index)
        for canonical in REQUIRED_COLUMNS:
            source_column = detected.get(canonical)
            normalized[canonical] = raw[source_column] if source_column else pd.NA

        fallback_vendor = source_path.stem.replace("_", " ").replace(" BOQ", "").strip()
        normalized["vendor"] = normalized["vendor"].fillna(fallback_vendor)
        normalized["source_file"] = source_path.name
        normalized["worksheet"] = sheet_name

        for numeric_column in ("quantity", "unit_rate", "total_amount"):
            normalized[numeric_column] = pd.to_numeric(normalized[numeric_column], errors="coerce")

        missing_total = normalized["total_amount"].isna() & normalized["quantity"].notna() & normalized["unit_rate"].notna()
        normalized.loc[missing_total, "total_amount"] = normalized.loc[missing_total, "quantity"] * normalized.loc[missing_total, "unit_rate"]
        return normalized.reset_index(drop=True)

    @staticmethod
    def _coerce_paths(files: str | Path | Iterable[str | Path]) -> list[Path]:
        """Return a list of paths whether the input is scalar or iterable."""
        if isinstance(files, (str, Path)):
            return [Path(files)]
        return [Path(file) for file in files]

    @staticmethod
    def _stringify_header(value: object, fallback_index: int) -> str:
        """Convert a worksheet cell value into a safe dataframe column name."""
        if pd.isna(value):
            return f"Unnamed: {fallback_index}"
        return str(value).strip()


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
    return ExcelReader().load(path).drop(columns=["worksheet"], errors="ignore")


def read_all_excels(directory: str | Path = "uploads") -> pd.DataFrame:
    """Read all Excel files from a directory into one master dataframe."""
    paths = sorted(Path(directory).glob("*.xlsx"))
    if not paths:
        paths = create_sample_files(directory)
    return ExcelReader().load(paths).drop(columns=["worksheet"], errors="ignore")
