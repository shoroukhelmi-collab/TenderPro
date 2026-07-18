"""Streamlit dashboard components for TenderPro."""

from __future__ import annotations

import streamlit as st

from modules.comparison_engine import build_comparison, supplier_ranking


def render_dashboard(normalized_data):
    """Render the MVP comparison, missing prices, and supplier ranking."""
    comparison = build_comparison(normalized_data)
    ranking = supplier_ranking(normalized_data)
    missing_prices = int(ranking["missing_prices"].sum()) if not ranking.empty else 0

    metric_cols = st.columns(3)
    metric_cols[0].metric("Suppliers", normalized_data["supplier"].nunique() if not normalized_data.empty else 0)
    metric_cols[1].metric("Compared Items", len(comparison))
    metric_cols[2].metric("Missing Prices", missing_prices)

    st.subheader("Item-by-item comparison")
    st.dataframe(comparison, use_container_width=True, hide_index=True)

    st.subheader("Supplier ranking")
    st.dataframe(ranking, use_container_width=True, hide_index=True)

    with st.expander("Normalized uploaded rows"):
        st.dataframe(normalized_data, use_container_width=True, hide_index=True)
