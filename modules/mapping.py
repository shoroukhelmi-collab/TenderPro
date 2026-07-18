"""Column detection and normalization helpers for TenderPro."""

from __future__ import annotations

import re
from typing import Iterable


COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "item_no": ("item no", "item", "item number", "item code", "code", "boq item", "ref", "reference"),
    "description": ("description", "item description", "desc", "scope", "work description", "particulars"),
    "unit": ("unit", "uom", "measure", "unit of measure", "units"),
    "quantity": ("quantity", "qty", "q ty", "quant", "boq qty"),
    "unit_rate": ("unit rate", "rate", "price", "unit price", "quoted rate", "supplier rate"),
    "total_amount": ("total amount", "amount", "total", "line total", "extended amount", "quoted amount"),
}

REQUIRED_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount")


def clean_column_name(value: object) -> str:
    """Return a lowercase, punctuation-insensitive column name."""
    cleaned = re.sub(r"[^a-zA-Z0-9]+", " ", str(value).strip().lower())
    return re.sub(r"\s+", " ", cleaned).strip()


def detect_columns(columns: Iterable[object]) -> dict[str, str]:
    """Detect common BOQ columns from a spreadsheet header row."""
    cleaned_columns = {clean_column_name(column): str(column) for column in columns}
    detected: dict[str, str] = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        cleaned_aliases = {clean_column_name(alias) for alias in aliases}
        for cleaned_column, original_column in cleaned_columns.items():
            if cleaned_column in cleaned_aliases:
                detected[canonical] = original_column
                break

    return detected
