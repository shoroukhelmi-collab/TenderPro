"""Core BOQ comparison calculations for TenderPro."""

from __future__ import annotations

import re

import pandas as pd

CANONICAL_COLUMNS = ("item_no", "description", "unit", "quantity", "unit_rate", "total_amount", "package", "supplier")
OUTLIER_RATIO_THRESHOLD = 1.5
OUTLIER_MIN_PRICED_SUPPLIERS = 3


def prepare_boq(data: pd.DataFrame) -> pd.DataFrame:
    """Return normalized rows with calculated totals and match keys."""
    frame = data.copy() if data is not None else pd.DataFrame()
    for column in CANONICAL_COLUMNS:
        if column not in frame.columns:
            frame[column] = pd.NA

    frame["supplier"] = frame["supplier"].fillna("Supplier").astype(str).str.strip().replace("", "Supplier")
    for column in ("quantity", "unit_rate", "total_amount"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    missing_total = frame["total_amount"].isna() & frame["quantity"].notna() & frame["unit_rate"].notna()
    frame.loc[missing_total, "total_amount"] = frame.loc[missing_total, "quantity"] * frame.loc[missing_total, "unit_rate"]
    frame = frame.dropna(subset=["item_no", "description"], how="all").reset_index(drop=True)
    frame["match_key"] = frame.apply(_match_key, axis=1)
    frame["package"] = frame["package"].fillna("").astype(str).str.strip()
    missing_package = frame["package"].eq("")
    frame.loc[missing_package, "package"] = frame.loc[missing_package].apply(_package_key, axis=1)
    frame["missing_price"] = frame["unit_rate"].isna() | frame["total_amount"].isna()
    return frame


def build_comparison(data: pd.DataFrame, excluded_outliers: set[str] | None = None) -> pd.DataFrame:
    """Create a master comparison table with lowest prices, variance, missing prices, and outlier flags."""
    prepared = prepare_boq(data)
    if prepared.empty:
        return pd.DataFrame()

    outliers = detect_outliers(prepared)
    outlier_keys = set(outliers["outlier_key"].dropna()) if not outliers.empty else set()
    excluded_outliers = excluded_outliers or set()
    prepared["outlier_key"] = prepared["match_key"] + "|" + prepared["supplier"]
    prepared["is_outlier"] = prepared["outlier_key"].isin(outlier_keys)
    prepared["excluded_outlier"] = prepared["outlier_key"].isin(excluded_outliers)

    valid = prepared[prepared["total_amount"].notna() & ~prepared["excluded_outlier"]]
    lowest_rows = valid.loc[valid.groupby("match_key")["total_amount"].idxmin()] if not valid.empty else pd.DataFrame()
    lowest_rows = lowest_rows[["match_key", "supplier", "total_amount"]].rename(
        columns={"supplier": "lowest_supplier", "total_amount": "lowest_price"}
    )

    grouped = prepared.groupby("match_key", dropna=False)
    summary = grouped.agg(
        item_no=("item_no", "first"),
        description=("description", "first"),
        unit=("unit", "first"),
        quantity=("quantity", "first"),
        package=("package", "first"),
        missing_prices=("missing_price", "sum"),
        outlier_count=("is_outlier", "sum"),
        excluded_outliers=("excluded_outlier", "sum"),
    ).reset_index()
    summary = summary.merge(lowest_rows, on="match_key", how="left")

    stats = valid.groupby("match_key")["total_amount"].agg(highest_price="max", average_price="mean").reset_index()
    summary = summary.merge(stats, on="match_key", how="left")
    summary["variance_amount"] = summary["highest_price"] - summary["lowest_price"]
    summary["variance_percent"] = (summary["variance_amount"] / summary["lowest_price"]).where(summary["lowest_price"].ne(0))

    suppliers = sorted(prepared["supplier"].dropna().unique().tolist())
    totals = prepared.pivot_table(index="match_key", columns="supplier", values="total_amount", aggfunc="first", dropna=False)
    totals = totals.reindex(columns=suppliers).reset_index()
    totals.columns.name = None
    totals = totals.rename(columns={column: f"{column} price" for column in totals.columns if column != "match_key"})

    flags = prepared.pivot_table(index="match_key", columns="supplier", values="is_outlier", aggfunc="max", dropna=False)
    flags = flags.reindex(columns=suppliers).reset_index()
    flags.columns.name = None
    flags = flags.rename(columns={column: f"{column} outlier" for column in flags.columns if column != "match_key"})

    result = summary.merge(totals, on="match_key", how="left").merge(flags, on="match_key", how="left")
    return result.drop(columns="match_key").sort_values(["package", "item_no", "description"], na_position="last").reset_index(drop=True)


def supplier_ranking(data: pd.DataFrame, excluded_outliers: set[str] | None = None) -> pd.DataFrame:
    """Rank suppliers by total priced amount, then by fewer missing prices."""
    prepared = prepare_boq(data)
    if prepared.empty:
        return pd.DataFrame(columns=["rank", "supplier", "quoted_total", "priced_items", "missing_prices", "lowest_items"])

    if excluded_outliers:
        prepared["outlier_key"] = prepared["match_key"] + "|" + prepared["supplier"]
        prepared.loc[prepared["outlier_key"].isin(excluded_outliers), "total_amount"] = pd.NA
    comparison = build_comparison(prepared, excluded_outliers)
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


def package_summary(data: pd.DataFrame, excluded_outliers: set[str] | None = None) -> pd.DataFrame:
    """Summarize totals by inferred package and supplier."""
    prepared = prepare_boq(data)
    if prepared.empty:
        return pd.DataFrame()
    if excluded_outliers:
        prepared["outlier_key"] = prepared["match_key"] + "|" + prepared["supplier"]
        prepared = prepared[~prepared["outlier_key"].isin(excluded_outliers)]
    return prepared.pivot_table(index="package", columns="supplier", values="total_amount", aggfunc="sum", fill_value=0).reset_index()


def missing_items(data: pd.DataFrame) -> pd.DataFrame:
    """Return rows with missing unit rates or totals."""
    prepared = prepare_boq(data)
    return prepared[prepared["missing_price"]][["supplier", "package", "item_no", "description", "unit", "quantity", "unit_rate", "total_amount"]]


def detect_outliers(data: pd.DataFrame) -> pd.DataFrame:
    """Flag supplier item prices that are unusually high versus the item median."""
    prepared = prepare_boq(data)
    priced = prepared[prepared["total_amount"].notna()].copy()
    if priced.empty:
        return pd.DataFrame(columns=["outlier_key"])
    stats = priced.groupby("match_key")["total_amount"].agg(median_price="median", priced_suppliers="count").reset_index()
    priced = priced.merge(stats, on="match_key", how="left")
    priced["variance_to_median_percent"] = (priced["total_amount"] - priced["median_price"]) / priced["median_price"].where(priced["median_price"].ne(0))
    mask = (priced["priced_suppliers"] >= OUTLIER_MIN_PRICED_SUPPLIERS) & (priced["total_amount"] > priced["median_price"] * OUTLIER_RATIO_THRESHOLD)
    outliers = priced[mask].copy()
    outliers["outlier_key"] = outliers["match_key"] + "|" + outliers["supplier"]
    return outliers[["outlier_key", "supplier", "package", "item_no", "description", "unit", "quantity", "unit_rate", "total_amount", "median_price", "variance_to_median_percent"]]


def dashboard_metrics(data: pd.DataFrame, uploaded_file_count: int = 0, excluded_outliers: set[str] | None = None) -> dict[str, object]:
    """Calculate high-level workflow KPIs."""
    prepared = prepare_boq(data)
    comparison = build_comparison(prepared, excluded_outliers) if not prepared.empty else pd.DataFrame()
    ranking = supplier_ranking(prepared, excluded_outliers) if not prepared.empty else pd.DataFrame()
    outliers = detect_outliers(prepared)
    lowest_total = float(ranking.iloc[0]["quoted_total"]) if not ranking.empty else 0.0
    highest_total = float(ranking["quoted_total"].max()) if not ranking.empty else 0.0
    return {
        "total_uploaded_files": uploaded_file_count,
        "detected_suppliers": int(prepared["supplier"].nunique()) if not prepared.empty else 0,
        "total_items": int(len(comparison)),
        "missing_prices": int(prepared["missing_price"].sum()) if not prepared.empty else 0,
        "outlier_items": int(len(outliers)),
        "estimated_saving": max(highest_total - lowest_total, 0.0),
        "lowest_total_supplier": ranking.iloc[0]["supplier"] if not ranking.empty else "—",
        "lowest_total_offer": lowest_total,
        "fully_priced_percent": 0 if prepared.empty else round(100 * (1 - prepared["missing_price"].mean()), 1),
    }


def _package_key(row: pd.Series) -> str:
    item_no = _clean(row.get("item_no"))
    if item_no:
        return re.split(r"[.\-\s]", item_no, maxsplit=1)[0] or "Unpackaged"
    return "Unpackaged"


def _match_key(row: pd.Series) -> str:
    item_no = _clean(row.get("item_no"))
    description = _clean(row.get("description"))
    unit = _clean(row.get("unit"))
    return f"item:{item_no}|unit:{unit}" if item_no else f"desc:{description}|unit:{unit}"


def _clean(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().casefold().split())
