"""Excel export helpers for TenderPro."""

from __future__ import annotations

from io import BytesIO

import pandas as pd

from modules.comparison_engine import build_comparison, dashboard_metrics, detect_outliers, missing_items, package_summary, supplier_ranking


def export_comparison(normalized_data: pd.DataFrame, uploaded_file_count: int = 0, excluded_outliers: set[str] | None = None) -> bytes:
    """Return an Excel workbook containing dashboard, comparison, summaries, pivots, and source data."""
    output = BytesIO()
    metrics = pd.DataFrame([dashboard_metrics(normalized_data, uploaded_file_count, excluded_outliers)])
    master = build_comparison(normalized_data, excluded_outliers)
    suppliers = supplier_ranking(normalized_data, excluded_outliers)
    packages = package_summary(normalized_data, excluded_outliers)
    missing = missing_items(normalized_data)
    outliers = detect_outliers(normalized_data)
    pivot = pd.DataFrame()
    if not normalized_data.empty:
        pivot = normalized_data.pivot_table(index="supplier", values=["quantity", "total_amount"], aggfunc={"quantity": "count", "total_amount": "sum"}).reset_index()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        metrics.to_excel(writer, sheet_name="Dashboard", index=False)
        master.to_excel(writer, sheet_name="Master Query", index=False)
        suppliers.to_excel(writer, sheet_name="Supplier Summary", index=False)
        packages.to_excel(writer, sheet_name="Package Summary", index=False)
        pivot.to_excel(writer, sheet_name="Pivot Summaries", index=False)
        missing.to_excel(writer, sheet_name="Missing Items", index=False)
        outliers.to_excel(writer, sheet_name="Outliers", index=False)
        normalized_data.to_excel(writer, sheet_name="Raw Data", index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)

    return output.getvalue()
