"""Column detection and normalization helpers for TenderPro."""

from __future__ import annotations

import re
from typing import Iterable


COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "item_no": ("item no", "item", "item number", "item code", "item no code", "item number code", "item item no code", "item no item code", "code", "boq item", "ref", "reference", "no", "no", "sl no", "s no", "serial no"),
    "description": ("description", "item description", "desc", "scope", "work description", "particulars", "item description of works"),
    "unit": ("unit", "uom", "measure", "unit of measure", "units", "u o m"),
    "quantity": ("quantity", "qty", "q ty", "quant", "boq qty", "qty boq"),
    "unit_rate": ("unit rate", "rate", "price", "unit price", "quoted rate", "supplier rate", "rate amount", "rate usd", "rate aed", "rate egp", "unit rate aed", "unit cost", "price usd", "price aed"),
    "total_amount": ("total amount", "amount", "total", "line total", "extended amount", "quoted amount", "total price", "total cost", "amount usd", "amount aed", "amount egp", "total amount aed"),
}

NEGATIVE_PRICE_HEADERS = {"rate only", "exchange rate", "discount rate", "tax rate", "vat rate"}

REQUIRED_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount")


def clean_column_name(value: object) -> str:
    """Return a lowercase, punctuation-insensitive column name."""
    cleaned = re.sub(r"[^a-zA-Z0-9]+", " ", str(value).strip().lower())
    return re.sub(r"\s+", " ", cleaned).strip()


def detect_columns(columns: Iterable[object]) -> dict[str, str]:
    """Detect common BOQ columns from a spreadsheet header row.

    Commercial BOQ exports often add currency, punctuation, or parenthetical text
    around rate/amount headers. Detection therefore accepts exact aliases first,
    then conservative token/substring matches for generic files.
    """
    cleaned_columns = {clean_column_name(column): str(column) for column in columns if clean_column_name(column)}
    detected: dict[str, str] = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        cleaned_aliases = {clean_column_name(alias) for alias in aliases}
        for cleaned_column, original_column in cleaned_columns.items():
            if cleaned_column in cleaned_aliases:
                detected[canonical] = original_column
                break
        if canonical not in detected:
            for cleaned_column, original_column in cleaned_columns.items():
                if _matches_alias(cleaned_column, cleaned_aliases, canonical):
                    detected[canonical] = original_column
                    break

    return detected


def _matches_alias(cleaned_column: str, aliases: set[str], canonical: str) -> bool:
    if cleaned_column in NEGATIVE_PRICE_HEADERS:
        return False
    tokens = set(cleaned_column.split())
    if canonical == "item_no":
        return _matches_item_number_header(cleaned_column, tokens)
    if canonical == "unit_rate":
        return ("rate" in tokens or "price" in tokens or {"unit", "cost"}.issubset(tokens)) and "total" not in tokens
    if canonical == "total_amount":
        return "amount" in tokens or "total" in tokens or {"total", "price"}.issubset(tokens) or {"total", "cost"}.issubset(tokens)
    return any(alias and (cleaned_column.startswith(alias + " ") or alias in cleaned_column) for alias in aliases)


def _matches_item_number_header(cleaned_column: str, tokens: set[str]) -> bool:
    """Return True for generic item-number/code headers, not description headers."""
    descriptive_tokens = {"description", "desc", "particulars", "scope", "work", "works"}
    if tokens & descriptive_tokens:
        return False
    if cleaned_column in {"item", "code", "ref", "reference", "no", "number"}:
        return True
    if tokens == {"item", "no"} or tokens == {"item", "number"}:
        return True
    if {"item", "no", "code"}.issubset(tokens) or {"item", "number", "code"}.issubset(tokens):
        return True
    if tokens == {"item", "code"} or tokens == {"boq", "item"}:
        return True
    if tokens in ({"sl", "no"}, {"s", "no"}, {"serial", "no"}):
        return True
    return cleaned_column.startswith(("item no ", "item number ", "item code ", "boq item ", "serial no ", "sl no ", "s no "))
