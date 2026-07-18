"""Excel export helpers for TenderPro."""

from __future__ import annotations

from io import BytesIO

import pandas as pd

from modules.comparison_engine import build_comparison, supplier_ranking


def export_comparison(normalized_data: pd.DataFrame) -> bytes:
    """Return an Excel workbook containing the final comparison and ranking."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        build_comparison(normalized_data).to_excel(writer, sheet_name="Comparison", index=False)
        supplier_ranking(normalized_data).to_excel(writer, sheet_name="Supplier Ranking", index=False)
        normalized_data.to_excel(writer, sheet_name="Normalized Data", index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)

    return output.getvalue()
