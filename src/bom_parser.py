"""Parse steel structure member data into a Bill of Materials."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

import pandas as pd

# Standard AISC / steel shape designations
STEEL_SHAPE_PATTERN = re.compile(
    r"\b("
    r"W\s*\d+\s*[xX×]\s*[\d.]+|"
    r"HSS\s*\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?(?:\s*[xX×]\s*[\d/]+)?|"
    r"L\s*\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?(?:\s*[xX×]\s*[\d/]+)?|"
    r"C\s*\d+\s*[xX×]\s*[\d.]+|"
    r"WT\s*\d+\s*[xX×]\s*[\d.]+|"
    r"PL\s*[\d/]+(?:\s*[\"']?\s*[xX×]\s*[\d/]+)?|"
    r"MC\s*\d+\s*[xX×]\s*[\d.]+"
    r")\b",
    re.IGNORECASE,
)

# Member marks used on GA / fabrication drawings (e.g. B2, BR1, BP1)
DRAWING_MARK_PATTERN = re.compile(
    r"\b(BR\d+|BP\d+|B[2-9]|PB[1-9][A-Z]?|PB[A-Z])\b",
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
    "mark": ("mark", "member mark", "member", "item", "designation", "size"),
    "description": ("description", "desc", "member type", "type"),
    "quantity": ("qty", "quantity", "no.", "no", "count", "pcs"),
    "length": ("length", "len", "lg", "l (ft)", "length (ft)"),
    "grade": ("grade", "material", "mat", "steel grade"),
    "weight": ("weight", "wt", "mass", "kg", "lbs"),
}

MARK_DESCRIPTIONS = {
    "B": "Beam / column member",
    "BR": "Bracing member",
    "BP": "Base plate",
    "PB": "Plan bay / section reference",
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
            items.extend(self._parse_bom_tables(page.tables, page.page_number))
            items.extend(self._parse_drawing_marks(page.text, page.page_number))
            items.extend(self._parse_steel_shapes(page.text, page.page_number))

        if not items:
            return self._empty_dataframe()

        df = pd.DataFrame([item.__dict__ for item in items])
        df = self._normalize_columns(df)
        df = self._aggregate_duplicates(df)
        return df

    def _parse_bom_tables(
        self, tables: list[list[list[str | None]]], page_number: int
    ) -> list[BOMItem]:
        items: list[BOMItem] = []
        for table in tables:
            if not table or len(table) < 2:
                continue
            header_map = self._map_table_headers(table[0])
            if not self._is_bom_table(header_map, table[0]):
                continue
            for row in table[1:]:
                if not row or all(not cell or not str(cell).strip() for cell in row):
                    continue
                item = self._row_to_item(row, header_map, page_number)
                if item:
                    items.append(item)
        return items

    @staticmethod
    def _is_bom_table(header_map: dict[str, int], header_row: list[str | None]) -> bool:
        if "mark" not in header_map:
            return False
        if len(header_map) < 2:
            return False
        max_header_len = max(
            (len(str(cell).strip()) for cell in header_row if cell),
            default=0,
        )
        return max_header_len <= 80

    def _map_table_headers(self, header_row: list[str | None]) -> dict[str, int]:
        mapping: dict[str, int] = {}
        for index, cell in enumerate(header_row):
            if not cell:
                continue
            normalized = str(cell).strip().lower()
            if len(normalized) > 80:
                continue
            for field, keywords in TABLE_HEADER_KEYWORDS.items():
                if field in mapping:
                    continue
                if any(
                    re.search(rf"\b{re.escape(keyword)}\b", normalized)
                    for keyword in keywords
                ):
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

        mark = cell("mark")
        if not mark or len(mark) > 40:
            return None

        quantity = self._parse_int(cell("quantity"), default=1)
        length = self._parse_float(cell("length"))
        weight = self._parse_float(cell("weight"))

        return BOMItem(
            mark=self._normalize_mark(mark),
            description=cell("description") or self._describe_mark(mark),
            quantity=quantity,
            length=length,
            grade=cell("grade").upper(),
            weight=weight,
            source_page=page_number,
        )

    def _parse_drawing_marks(self, text: str, page_number: int) -> list[BOMItem]:
        if not text.strip():
            return []

        counts = Counter(
            self._normalize_mark(match)
            for match in DRAWING_MARK_PATTERN.findall(text)
        )
        return [
            BOMItem(
                mark=mark,
                description=self._describe_mark(mark),
                quantity=count,
                source_page=page_number,
            )
            for mark, count in sorted(counts.items())
        ]

    def _parse_steel_shapes(self, text: str, page_number: int) -> list[BOMItem]:
        if not text.strip():
            return []

        items: list[BOMItem] = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            matches = STEEL_SHAPE_PATTERN.findall(line)
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

    @classmethod
    def _describe_mark(cls, mark: str) -> str:
        normalized = cls._normalize_mark(mark)
        if normalized.startswith("BR"):
            return MARK_DESCRIPTIONS["BR"]
        if normalized.startswith("BP"):
            return MARK_DESCRIPTIONS["BP"]
        if normalized.startswith("PB"):
            return MARK_DESCRIPTIONS["PB"]
        if re.fullmatch(r"B[2-9]", normalized):
            return MARK_DESCRIPTIONS["B"]
        return "Structural member"

    @staticmethod
    def _normalize_mark(mark: str) -> str:
        normalized = re.sub(r"\s+", "", mark.upper())
        return normalized.replace("×", "X")

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
            .sort_values(["mark"], na_position="last")
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
