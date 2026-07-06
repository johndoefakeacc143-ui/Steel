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
        summary = self.build_summary(df)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="BOM", index=False)
            self._set_widths(writer.sheets["BOM"], df.columns)
            summary.to_excel(writer, sheet_name="Summary", index=False)
            self._set_widths(writer.sheets["Summary"], summary.columns)
        return output_path

    @staticmethod
    def build_summary(df: pd.DataFrame) -> pd.DataFrame:
        """Totals per description/section: quantity, length and weight."""
        columns = ["description", "total_quantity", "total_length", "total_weight"]
        if df.empty:
            return pd.DataFrame(columns=columns)

        work = df.copy()
        qty = pd.to_numeric(work.get("quantity"), errors="coerce").fillna(0)
        length = pd.to_numeric(work.get("length"), errors="coerce").fillna(0)
        weight = pd.to_numeric(work.get("weight"), errors="coerce").fillna(0)
        work["_qty"] = qty
        work["_len_total"] = length * qty
        work["_wt_total"] = weight * qty
        work["description"] = work.get("description", "").fillna("").replace("", "NA")

        grouped = (
            work.groupby("description", as_index=False)
            .agg(total_quantity=("_qty", "sum"), total_length=("_len_total", "sum"), total_weight=("_wt_total", "sum"))
            .sort_values("description")
        )
        grouped["total_quantity"] = grouped["total_quantity"].astype(int)
        grouped["total_length"] = grouped["total_length"].round(2)
        grouped["total_weight"] = grouped["total_weight"].round(2)

        total_row = pd.DataFrame(
            [{
                "description": "TOTAL",
                "total_quantity": int(grouped["total_quantity"].sum()),
                "total_length": round(float(grouped["total_length"].sum()), 2),
                "total_weight": round(float(grouped["total_weight"].sum()), 2),
            }]
        )
        return pd.concat([grouped, total_row], ignore_index=True)[columns]

    def _set_widths(self, worksheet, columns) -> None:
        for column_index, column_name in enumerate(columns, start=1):
            width = self.COLUMN_WIDTHS.get(column_name, 18)
            worksheet.column_dimensions[self._column_letter(column_index)].width = width

    @staticmethod
    def _column_letter(index: int) -> str:
        result = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result
