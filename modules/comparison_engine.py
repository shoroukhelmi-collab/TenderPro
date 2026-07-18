"""Generic commercial comparison calculations for TenderPro."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

import pandas as pd


DEFAULT_MATCHING_RULES = ("item_no", "description", "unit")
CANONICAL_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount", "vendor")


@dataclass(frozen=True)
class ComparisonEngine:
    """Build supplier-agnostic tender comparison tables from normalized BOQ rows.

    The engine expects the normalized dataframe returned by ``ExcelReader`` helpers:
    one row per supplier quotation line using TenderPro's canonical column names.
    Matching is configurable and defaults to item number, description, and unit. Rows
    without an item number automatically match on description and unit instead.
    """

    matching_rules: Iterable[str] = field(default_factory=lambda: DEFAULT_MATCHING_RULES)
    supplier_column: str = "vendor"

    def compare(self, normalized_data: pd.DataFrame) -> pd.DataFrame:
        """Return one master comparison table supporting any number of suppliers."""
        prepared = self._prepare(normalized_data)
        if prepared.empty:
            return self._empty_comparison()

        prepared["match_key"] = prepared.apply(self._build_match_key, axis=1)
        prepared["is_missing_price"] = prepared["unit_rate"].isna() | prepared["total_amount"].isna()

        summary = self._line_summary(prepared)
        supplier_columns = self._supplier_price_columns(prepared)

        master = summary.merge(supplier_columns, on="match_key", how="left")
        master = master.sort_values(["item_no", "description", "unit"], na_position="last").reset_index(drop=True)
        return master.drop(columns=["match_key"])

    def enrich_rows(self, normalized_data: pd.DataFrame) -> pd.DataFrame:
        """Return row-level comparison details for backwards-compatible dashboards."""
        prepared = self._prepare(normalized_data)
        if prepared.empty:
            return prepared

        prepared["line_key"] = prepared.apply(self._build_match_key, axis=1)
        prepared["is_missing_price"] = prepared["unit_rate"].isna() | prepared["total_amount"].isna()
        prepared["lowest_total"] = prepared.groupby("line_key")["total_amount"].transform("min")
        prepared["lowest_unit_rate"] = prepared.groupby("line_key")["unit_rate"].transform("min")
        prepared["is_lowest_offer"] = prepared["total_amount"].eq(prepared["lowest_total"]) & prepared["total_amount"].notna()
        prepared["variance_amount"] = prepared["total_amount"] - prepared["lowest_total"]
        prepared["variance_percent"] = self._variance_percent(prepared["total_amount"], prepared["lowest_total"])
        return prepared

    def detect_suppliers(self, normalized_data: pd.DataFrame) -> list[str]:
        """Return suppliers present in the normalized data without fixed limits."""
        if self.supplier_column not in normalized_data.columns:
            return []
        return sorted(normalized_data[self.supplier_column].dropna().astype(str).unique().tolist())

    def _prepare(self, data: pd.DataFrame) -> pd.DataFrame:
        frame = data.copy() if data is not None else pd.DataFrame()
        for column in CANONICAL_COLUMNS:
            if column not in frame.columns:
                frame[column] = pd.NA
        frame[self.supplier_column] = frame[self.supplier_column].fillna("Unspecified Supplier").astype(str).str.strip()
        for column in ("quantity", "unit_rate", "total_amount"):
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
        missing_total = frame["total_amount"].isna() & frame["quantity"].notna() & frame["unit_rate"].notna()
        frame.loc[missing_total, "total_amount"] = frame.loc[missing_total, "quantity"] * frame.loc[missing_total, "unit_rate"]
        return frame.dropna(subset=["item_no", "description", "unit"], how="all")

    def _build_match_key(self, row: pd.Series) -> tuple[tuple[str, str], ...]:
        rules = tuple(self.matching_rules) or DEFAULT_MATCHING_RULES
        item_no = self._clean_value(row.get("item_no"))
        active_rules = tuple(rule for rule in rules if rule != "item_no") if not item_no else rules
        return tuple((rule, self._clean_value(row.get(rule))) for rule in active_rules)

    @staticmethod
    def _clean_value(value: object) -> str:
        if pd.isna(value):
            return ""
        return " ".join(str(value).strip().casefold().split())

    def _line_summary(self, prepared: pd.DataFrame) -> pd.DataFrame:
        priced = prepared[prepared["total_amount"].notna()].copy()
        lowest_idx = priced.groupby("match_key")["total_amount"].idxmin() if not priced.empty else pd.Index([])
        lowest_supplier = priced.loc[lowest_idx, ["match_key", self.supplier_column]].rename(columns={self.supplier_column: "lowest_supplier"})

        summary = (
            prepared.groupby("match_key", dropna=False)
            .agg(
                item_no=("item_no", "first"),
                description=("description", "first"),
                unit=("unit", "first"),
                quantity=("quantity", "first"),
                lowest_unit_rate=("unit_rate", "min"),
                lowest_total_amount=("total_amount", "min"),
                missing_prices=("is_missing_price", "sum"),
                number_of_quotations=(self.supplier_column, "nunique"),
            )
            .reset_index()
        )
        summary = summary.merge(lowest_supplier, on="match_key", how="left")
        summary["variance_percent"] = self._group_variance(prepared)
        return summary

    def _supplier_price_columns(self, prepared: pd.DataFrame) -> pd.DataFrame:
        pivot = prepared.pivot_table(index="match_key", columns=self.supplier_column, values="total_amount", aggfunc="first").reset_index()
        pivot.columns.name = None
        return pivot.rename(columns={column: f"{column} total_amount" for column in pivot.columns if column != "match_key"})

    @staticmethod
    def _variance_percent(current: pd.Series, lowest: pd.Series) -> pd.Series:
        variance = pd.Series(pd.NA, index=current.index, dtype="Float64")
        mask = lowest.notna() & lowest.ne(0) & current.notna()
        variance.loc[mask] = (current.loc[mask] - lowest.loc[mask]) / lowest.loc[mask] * 100
        return variance

    @staticmethod
    def _group_variance(prepared: pd.DataFrame) -> pd.Series:
        totals = prepared.groupby("match_key")["total_amount"]
        lowest = totals.min()
        highest = totals.max()
        variance = ((highest - lowest) / lowest * 100).where(lowest.notna() & lowest.ne(0), pd.NA)
        return variance.reset_index(drop=True)

    @staticmethod
    def _empty_comparison() -> pd.DataFrame:
        return pd.DataFrame(
            columns=[
                "item_no", "description", "unit", "quantity", "lowest_unit_rate", "lowest_total_amount",
                "lowest_supplier", "missing_prices", "variance_percent", "number_of_quotations",
            ]
        )


def enrich_comparison(master: pd.DataFrame) -> pd.DataFrame:
    """Add lowest-offer, variance, and missing-price flags to normalized BOQ rows."""
    return ComparisonEngine().enrich_rows(master)


def generate_master_comparison(master: pd.DataFrame, matching_rules: Iterable[str] | None = None) -> pd.DataFrame:
    """Create the generic supplier comparison table."""
    return ComparisonEngine(matching_rules or DEFAULT_MATCHING_RULES).compare(master)


def vendor_ranking(comparison: pd.DataFrame) -> pd.DataFrame:
    """Rank suppliers by total quoted amount and price coverage."""
    if comparison.empty:
        return pd.DataFrame(columns=["vendor", "quoted_total", "priced_items", "missing_prices", "lowest_offers", "rank"])
    ranking = comparison.groupby("vendor", dropna=False).agg(
        quoted_total=("total_amount", "sum"), priced_items=("total_amount", "count"),
        missing_prices=("is_missing_price", "sum"), lowest_offers=("is_lowest_offer", "sum"),
    ).reset_index().sort_values(["quoted_total", "missing_prices"], ascending=[True, True])
    ranking["rank"] = range(1, len(ranking) + 1)
    return ranking


def comparison_pivot(comparison: pd.DataFrame) -> pd.DataFrame:
    """Create a side-by-side item comparison table by supplier."""
    if comparison.empty:
        return pd.DataFrame()
    return generate_master_comparison(comparison)
