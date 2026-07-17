"""Export helpers for TenderPro comparison workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from modules.comparison_engine import comparison_pivot, vendor_ranking


def export_comparison(comparison: pd.DataFrame, directory: str | Path = "exports") -> Path:
    """Write the current comparison to a formatted Excel workbook."""
    export_dir = Path(directory)
    export_dir.mkdir(parents=True, exist_ok=True)
    output_path = export_dir / "tenderpro_comparison.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        comparison.to_excel(writer, sheet_name="Normalized Data", index=False)
        comparison_pivot(comparison).to_excel(writer, sheet_name="Side by Side", index=False)
        vendor_ranking(comparison).to_excel(writer, sheet_name="Vendor Ranking", index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)

    return output_path
