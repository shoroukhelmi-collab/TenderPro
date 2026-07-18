"""Streamlit dashboard components for TenderPro."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from modules.comparison_engine import (
    build_comparison,
    dashboard_metrics,
    detect_outliers,
    missing_items,
    package_summary,
    supplier_ranking,
)


def render_home_kpis(normalized_data: pd.DataFrame | None, uploaded_file_count: int = 0) -> None:
    """Render simple home-page KPI cards for the current comparison."""
    metrics = dashboard_metrics(normalized_data if normalized_data is not None else pd.DataFrame(), uploaded_file_count)
    cols = st.columns(6)
    labels = [
        ("Total uploaded files", metrics["total_uploaded_files"]),
        ("Detected suppliers", metrics["detected_suppliers"]),
        ("Total items", metrics["total_items"]),
        ("Missing prices", metrics["missing_prices"]),
        ("Outlier items", metrics["outlier_items"]),
        ("Estimated saving", f"{metrics['estimated_saving']:,.2f}"),
    ]
    for column, (label, value) in zip(cols, labels, strict=False):
        column.metric(label, value)


def render_dashboard(normalized_data: pd.DataFrame, uploaded_file_count: int = 0):
    """Render dashboard, master query, summaries, missing prices, and outlier controls."""
    excluded = st.session_state.setdefault("excluded_outliers", set())
    comparison = build_comparison(normalized_data, excluded)
    ranking = supplier_ranking(normalized_data, excluded)
    packages = package_summary(normalized_data, excluded)
    missing = missing_items(normalized_data)
    outliers = detect_outliers(normalized_data)
    metrics = dashboard_metrics(normalized_data, uploaded_file_count, excluded)

    tabs = st.tabs(["Dashboard", "Master Query", "Supplier Summary", "Package Summary", "Missing Items", "Outliers"])
    with tabs[0]:
        _render_main_dashboard(metrics, ranking, packages, comparison)
    with tabs[1]:
        st.subheader("Master Query")
        st.caption("Outlier columns show detected unusual prices; excluded outliers stay visible but are ignored in totals.")
        st.dataframe(comparison, use_container_width=True, hide_index=True)
    with tabs[2]:
        st.subheader("Supplier Summary")
        st.dataframe(ranking, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.subheader("Package Summary")
        st.dataframe(packages, use_container_width=True, hide_index=True)
    with tabs[4]:
        st.subheader("Missing Items")
        st.dataframe(missing, use_container_width=True, hide_index=True)
    with tabs[5]:
        _render_outliers(outliers, excluded)


def _render_main_dashboard(metrics: dict[str, object], ranking: pd.DataFrame, packages: pd.DataFrame, comparison: pd.DataFrame) -> None:
    top = st.columns(5)
    top[0].metric("Lowest total supplier", metrics["lowest_total_supplier"])
    top[1].metric("Lowest total offer", f"{metrics['lowest_total_offer']:,.2f}")
    top[2].metric("Potential saving", f"{metrics['estimated_saving']:,.2f}")
    top[3].metric("Missing price count", metrics["missing_prices"])
    top[4].metric("Outlier count", metrics["outlier_items"])
    st.metric("Fully priced items percentage", f"{metrics['fully_priced_percent']}%")

    chart_cols = st.columns(3)
    if not ranking.empty:
        chart_cols[0].bar_chart(ranking.set_index("supplier")["quoted_total"])
    if not packages.empty:
        lowest_by_package = packages.set_index("package").idxmin(axis=1).value_counts()
        chart_cols[1].bar_chart(lowest_by_package)
    if not comparison.empty and "variance_amount" in comparison:
        high_variance = comparison.nlargest(10, "variance_amount").set_index("description")["variance_amount"]
        chart_cols[2].bar_chart(high_variance)


def _render_outliers(outliers: pd.DataFrame, excluded: set[str]) -> None:
    st.subheader("Outliers")
    if outliers.empty:
        st.info("No price outliers were detected.")
        return
    edited = outliers.copy()
    edited["exclude_from_totals"] = edited["outlier_key"].isin(excluded)
    edited = st.data_editor(
        edited,
        use_container_width=True,
        hide_index=True,
        disabled=[column for column in edited.columns if column != "exclude_from_totals"],
        key="outlier_editor",
    )
    st.session_state["excluded_outliers"] = set(edited.loc[edited["exclude_from_totals"], "outlier_key"].dropna())
    st.caption("Excluding an outlier affects summaries and export calculations only; raw source rows are preserved.")
