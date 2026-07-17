"""Streamlit dashboard components for TenderPro."""

from __future__ import annotations

import streamlit as st

from modules.comparison_engine import comparison_pivot, vendor_ranking


def render_dashboard(comparison):
    """Render KPI cards and comparison tables."""
    ranking = vendor_ranking(comparison)
    missing_prices = int(comparison["is_missing_price"].sum()) if not comparison.empty else 0
    lowest_count = int(comparison["is_lowest_offer"].sum()) if not comparison.empty else 0
    best_vendor = ranking.iloc[0]["vendor"] if not ranking.empty else "N/A"

    metric_cols = st.columns(4)
    metric_cols[0].metric("Best Overall Offer", best_vendor)
    metric_cols[1].metric("Lowest Line Offers", lowest_count)
    metric_cols[2].metric("Missing Prices", missing_prices)
    metric_cols[3].metric("Vendors", comparison["vendor"].nunique() if not comparison.empty else 0)

    st.subheader("Vendor Ranking")
    st.dataframe(ranking, use_container_width=True, hide_index=True)

    st.subheader("Side-by-Side BOQ Comparison")
    st.dataframe(comparison_pivot(comparison), use_container_width=True, hide_index=True)

    st.subheader("Detailed Variance Analysis")
    st.dataframe(comparison, use_container_width=True, hide_index=True)
