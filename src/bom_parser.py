"""Parse steel structure member data into a Bill of Materials."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

# Common steel member designations found on structural drawings
MEMBER_PATTERN = re.compile(
    r"\b("
    r"W\s*\d+\s*[xX×]\s*[\d.]+|"  # Wide flange: W12x26
    r"HSS\s*\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?(?:\s*[xX×]\s*[\d/]+)?|"  # HSS
    r"L\s*\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?(?:\s*[xX×]\s*[\d/]+)?|"  # Angle
    r"C\s*\d+\s*[xX×]\s*[\d.]+|"  # Channel
    r"WT\s*\d+\s*[xX×]\s*[\d.]+|"  # Tee
    r"PL\s*[\d/]+(?:\s*[\"']?\s*[xX×]\s*[\d/]+)?|"  # Plate
    r"MC\s*\d+\s*[xX×]\s*[\d.]+"  # Misc channel
    r")\b",
    re.IGNORECASE,
)

LENGTH_PATTERN = re.compile(
    r"(?:length|len|lg|l)\s*[:=]?\s*(\d+(?:\.\d+)?)\s*(?:['\"]|ft|feet|in|inch|mm)?",
    re.IGNORECASE,
)

QTY_PATTERN = re.compile(
    r"(?:qty|quantity|no\.?|count)\s*[:=]?\s*(\d+)",
    re.IGNORECASE,
)

GRADE_PATTERN = re.compile(
    r"\b(A(?:36|572|992|500)|S275|S355|ST37|ST52)\b",
    re.IGNORECASE,
)

TABLE_HEADER_KEYWORDS = {
    "mark": ("mark", "member", "item", "designation", "size"),
    "description": ("description", "desc", "member type", "type"),
    "quantity": ("qty", "quantity", "no", "count", "pcs"),
    "length": ("length", "len", "lg", "l (ft)", "length (ft)"),
    "grade": ("grade", "material", "mat", "steel grade"),
    "weight": ("weight", "wt", "mass", "kg", "lbs"),
}


@dataclass
class BOMItem:
    mark: str
    description: str = ""
    quantity: int = 1
    length: float | None = None
    grade: str = ""
    weight: float | None = None
    source_page: int | None = None


class BOMParser:
    """Build a BOM DataFrame from extracted PDF content."""

    def parse_pages(self, pages: list) -> pd.DataFrame:
        items: list[BOMItem] = []

        for page in pages:
            items.extend(self._parse_tables(page.tables, page.page_number))
            items.extend(self._parse_text(page.text, page.page_number))

        if not items:
            return self._empty_dataframe()

        df = pd.DataFrame([item.__dict__ for item in items])
        df = self._normalize_columns(df)
        df = self._aggregate_duplicates(df)
        return df

    def _parse_tables(
        self, tables: list[list[list[str | None]]], page_number: int
    ) -> list[BOMItem]:
        items: list[BOMItem] = []
        for table in tables:
            if not table or len(table) < 2:
                continue
            header_map = self._map_table_headers(table[0])
            if not header_map:
                continue
            for row in table[1:]:
                if not row or all(not cell or not str(cell).strip() for cell in row):
                    continue
                item = self._row_to_item(row, header_map, page_number)
                if item:
                    items.append(item)
        return items

    def _map_table_headers(self, header_row: list[str | None]) -> dict[str, int]:
        mapping: dict[str, int] = {}
        for index, cell in enumerate(header_row):
            if not cell:
                continue
            normalized = str(cell).strip().lower()
            for field, keywords in TABLE_HEADER_KEYWORDS.items():
                if any(keyword in normalized for keyword in keywords):
                    mapping[field] = index
                    break
        return mapping

    def _row_to_item(
        self, row: list[str | None], header_map: dict[str, int], page_number: int
    ) -> BOMItem | None:
        def cell(field: str) -> str:
            index = header_map.get(field)
            if index is None or index >= len(row) or row[index] is None:
                return ""
            return str(row[index]).strip()

        mark = cell("mark") or cell("description")
        if not mark:
            return None

        quantity = self._parse_int(cell("quantity"), default=1)
        length = self._parse_float(cell("length"))
        weight = self._parse_float(cell("weight"))

        return BOMItem(
            mark=self._normalize_mark(mark),
            description=cell("description") or mark,
            quantity=quantity,
            length=length,
            grade=cell("grade").upper(),
            weight=weight,
            source_page=page_number,
        )

    def _parse_text(self, text: str, page_number: int) -> list[BOMItem]:
        if not text.strip():
            return []

        items: list[BOMItem] = []
        lines = [line.strip() for line in text.splitlines() if line.strip()]

        for line in lines:
            matches = MEMBER_PATTERN.findall(line)
            if not matches:
                continue

            grade_match = GRADE_PATTERN.search(line)
            length_match = LENGTH_PATTERN.search(line)
            qty_match = QTY_PATTERN.search(line)

            for match in matches:
                items.append(
                    BOMItem(
                        mark=self._normalize_mark(match),
                        description=match.strip(),
                        quantity=int(qty_match.group(1)) if qty_match else 1,
                        length=float(length_match.group(1)) if length_match else None,
                        grade=grade_match.group(1).upper() if grade_match else "",
                        source_page=page_number,
                    )
                )
        return items

    @staticmethod
    def _normalize_mark(mark: str) -> str:
        normalized = re.sub(r"\s+", "", mark.upper())
        normalized = normalized.replace("×", "X")
        return normalized

    @staticmethod
    def _parse_int(value: str, default: int = 1) -> int:
        try:
            return int(float(value))
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _parse_float(value: str) -> float | None:
        try:
            cleaned = re.sub(r"[^\d.]", "", value)
            return float(cleaned) if cleaned else None
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
        column_order = [
            "mark",
            "description",
            "quantity",
            "length",
            "grade",
            "weight",
            "source_page",
        ]
        for column in column_order:
            if column not in df.columns:
                df[column] = None
        return df[column_order]

    @staticmethod
    def _aggregate_duplicates(df: pd.DataFrame) -> pd.DataFrame:
        group_cols = ["mark", "description", "length", "grade", "source_page"]
        aggregated = (
            df.groupby(group_cols, dropna=False, as_index=False)
            .agg({"quantity": "sum", "weight": "first"})
            .sort_values(["mark", "source_page"], na_position="last")
            .reset_index(drop=True)
        )
        aggregated.insert(0, "item_no", range(1, len(aggregated) + 1))
        return aggregated

    @staticmethod
    def _empty_dataframe() -> pd.DataFrame:
        return pd.DataFrame(
            columns=[
                "item_no",
                "mark",
                "description",
                "quantity",
                "length",
                "grade",
                "weight",
                "source_page",
            ]
        )
