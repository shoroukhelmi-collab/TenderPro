"""Core BOQ comparison calculations for TenderPro."""

from __future__ import annotations

import pandas as pd

CANONICAL_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount", "supplier")


def prepare_boq(data: pd.DataFrame) -> pd.DataFrame:
    """Return normalized rows with calculated totals and match keys."""
    frame = data.copy() if data is not None else pd.DataFrame()
    for column in CANONICAL_COLUMNS:
        if column not in frame.columns:
            frame[column] = pd.NA

    frame["supplier"] = frame["supplier"].fillna("Supplier").astype(str).str.strip()
    for column in ("quantity", "unit_rate", "total_amount"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    missing_total = frame["total_amount"].isna() & frame["quantity"].notna() & frame["unit_rate"].notna()
    frame.loc[missing_total, "total_amount"] = frame.loc[missing_total, "quantity"] * frame.loc[missing_total, "unit_rate"]
    frame = frame.dropna(subset=["item_no", "description"], how="all").reset_index(drop=True)
    frame["match_key"] = frame.apply(_match_key, axis=1)
    frame["missing_price"] = frame["unit_rate"].isna() | frame["total_amount"].isna()
    return frame


def build_comparison(data: pd.DataFrame) -> pd.DataFrame:
    """Compare suppliers item by item and show lowest and missing prices."""
    prepared = prepare_boq(data)
    if prepared.empty:
        return pd.DataFrame(columns=["item_no", "description", "unit", "quantity", "lowest_supplier", "lowest_price", "missing_prices"])

    priced = prepared[prepared["total_amount"].notna()]
    lowest_rows = prepared.loc[priced.groupby("match_key")["total_amount"].idxmin()] if not priced.empty else pd.DataFrame()
    lowest_rows = lowest_rows[["match_key", "supplier", "total_amount"]].rename(
        columns={"supplier": "lowest_supplier", "total_amount": "lowest_price"}
    )

    summary = prepared.groupby("match_key", dropna=False).agg(
        item_no=("item_no", "first"),
        description=("description", "first"),
        unit=("unit", "first"),
        quantity=("quantity", "first"),
        missing_prices=("missing_price", "sum"),
    ).reset_index()
    summary = summary.merge(lowest_rows, on="match_key", how="left")

    suppliers = sorted(prepared["supplier"].dropna().unique().tolist())
    totals = prepared.pivot_table(index="match_key", columns="supplier", values="total_amount", aggfunc="first", dropna=False)
    totals = totals.reindex(columns=suppliers).reset_index()
    totals.columns.name = None
    totals = totals.rename(columns={column: f"{column} price" for column in totals.columns if column != "match_key"})

    result = summary.merge(totals, on="match_key", how="left")
    return result.drop(columns="match_key").sort_values(["item_no", "description"], na_position="last").reset_index(drop=True)


def supplier_ranking(data: pd.DataFrame) -> pd.DataFrame:
    """Rank suppliers by total priced amount, then by fewer missing prices."""
    prepared = prepare_boq(data)
    if prepared.empty:
        return pd.DataFrame(columns=["rank", "supplier", "quoted_total", "priced_items", "missing_prices", "lowest_items"])

    comparison = build_comparison(prepared)
    lowest_counts = comparison["lowest_supplier"].value_counts(dropna=True).rename_axis("supplier").reset_index(name="lowest_items")
    ranking = prepared.groupby("supplier", dropna=False).agg(
        quoted_total=("total_amount", "sum"),
        priced_items=("total_amount", "count"),
        missing_prices=("missing_price", "sum"),
    ).reset_index()
    ranking = ranking.merge(lowest_counts, on="supplier", how="left").fillna({"lowest_items": 0})
    ranking = ranking.sort_values(["quoted_total", "missing_prices", "supplier"], ascending=[True, True, True]).reset_index(drop=True)
    ranking["rank"] = ranking.index + 1
    return ranking[["rank", "supplier", "quoted_total", "priced_items", "missing_prices", "lowest_items"]]


def _match_key(row: pd.Series) -> str:
    item_no = _clean(row.get("item_no"))
    description = _clean(row.get("description"))
    unit = _clean(row.get("unit"))
    return f"item:{item_no}|unit:{unit}" if item_no else f"desc:{description}|unit:{unit}"


def _clean(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().casefold().split())
