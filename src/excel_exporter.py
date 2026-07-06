"""Export BOM DataFrames to Excel."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


class ExcelExporter:
    """Write Bill of Materials to an Excel workbook."""

    COLUMN_WIDTHS = {
        "item_no": 10,
        "mark": 18,
        "description": 30,
        "quantity": 12,
        "length": 12,
        "grade": 12,
        "weight": 12,
        "source_page": 14,
    }

    def export(self, df: pd.DataFrame, output_path: Path) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="BOM", index=False)
            worksheet = writer.sheets["BOM"]
            for column_index, column_name in enumerate(df.columns, start=1):
                width = self.COLUMN_WIDTHS.get(column_name, 15)
                col_letter = self._column_letter(column_index)
                worksheet.column_dimensions[col_letter].width = width
        return output_path

    @staticmethod
    def _column_letter(index: int) -> str:
        result = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result
