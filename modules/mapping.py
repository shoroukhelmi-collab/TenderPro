"""Column detection and normalization helpers for TenderPro."""

from __future__ import annotations

import re
from typing import Dict, Iterable


COLUMN_MAPPINGS: dict[str, tuple[str, ...]] = {
    "item_no": ("item no", "item", "code", "item number", "item code"),
    "description": ("description", "item description", "desc", "scope"),
    "unit": ("unit", "uom", "measure", "unit of measure"),
    "quantity": ("quantity", "qty", "q'ty", "quant"),
    "unit_rate": ("unit rate", "price", "rate", "unit price"),
    "total_amount": ("total amount", "amount", "total", "line total"),
    "vendor": ("vendor", "supplier", "bidder", "company"),
}

REQUIRED_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount", "vendor")


def _clean_name(value: object) -> str:
    """Return a lowercase, punctuation-insensitive column name."""
    cleaned = re.sub(r"[^a-zA-Z0-9]+", " ", str(value).strip().lower())
    return re.sub(r"\s+", " ", cleaned).strip()


def detect_columns(columns: Iterable[object]) -> Dict[str, str]:
    """Map vendor-specific columns to TenderPro's canonical schema."""
    cleaned_columns = {_clean_name(column): str(column) for column in columns}
    detected: dict[str, str] = {}

    for canonical, aliases in COLUMN_MAPPINGS.items():
        cleaned_aliases = {_clean_name(alias) for alias in aliases}
        for cleaned_column, original_column in cleaned_columns.items():
            if cleaned_column in cleaned_aliases:
                detected[canonical] = original_column
                break

    return detected
