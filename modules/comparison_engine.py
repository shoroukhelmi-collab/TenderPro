"""Commercial comparison calculations for TenderPro."""

from __future__ import annotations

import pandas as pd


def enrich_comparison(master: pd.DataFrame) -> pd.DataFrame:
    """Add lowest-offer, variance, and missing-price flags to normalized BOQ rows."""
    if master.empty:
        return master.copy()

    comparison = master.copy()
    comparison["line_key"] = comparison["item_no"].astype(str).str.strip() + " | " + comparison["description"].astype(str).str.strip()
    comparison["is_missing_price"] = comparison["unit_rate"].isna() | comparison["total_amount"].isna()
    comparison["lowest_total"] = comparison.groupby("line_key")["total_amount"].transform("min")
    comparison["is_lowest_offer"] = comparison["total_amount"].eq(comparison["lowest_total"]) & comparison["total_amount"].notna()
    comparison["variance_amount"] = comparison["total_amount"] - comparison["lowest_total"]
    comparison["variance_percent"] = pd.NA
    has_nonzero_lowest = comparison["lowest_total"].notna() & comparison["lowest_total"].ne(0)
    comparison.loc[has_nonzero_lowest, "variance_percent"] = (
        comparison.loc[has_nonzero_lowest, "variance_amount"] / comparison.loc[has_nonzero_lowest, "lowest_total"] * 100
    )
    return comparison


def vendor_ranking(comparison: pd.DataFrame) -> pd.DataFrame:
    """Rank vendors by total quoted amount and price coverage."""
    if comparison.empty:
        return pd.DataFrame(columns=["vendor", "quoted_total", "priced_items", "missing_prices", "lowest_offers", "rank"])

    ranking = (
        comparison.groupby("vendor", dropna=False)
        .agg(
            quoted_total=("total_amount", "sum"),
            priced_items=("total_amount", "count"),
            missing_prices=("is_missing_price", "sum"),
            lowest_offers=("is_lowest_offer", "sum"),
        )
        .reset_index()
        .sort_values(["quoted_total", "missing_prices"], ascending=[True, True])
    )
    ranking["rank"] = range(1, len(ranking) + 1)
    return ranking


def comparison_pivot(comparison: pd.DataFrame) -> pd.DataFrame:
    """Create a side-by-side item comparison table by vendor."""
    if comparison.empty:
        return pd.DataFrame()

    pivot = comparison.pivot_table(
        index=["item_no", "description", "unit", "quantity"],
        columns="vendor",
        values="total_amount",
        aggfunc="first",
    ).reset_index()
    pivot.columns.name = None
    return pivot
