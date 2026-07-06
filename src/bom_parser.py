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

# Linear members (beams/bracings) can carry a length; plates and plan-bay
# references cannot.
LINEAR_MARK_PATTERN = re.compile(r"^(B[2-9]|BR\d+)$", re.IGNORECASE)


def _is_linear_mark(mark: str) -> bool:
    return bool(LINEAR_MARK_PATTERN.match(mark))

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

    def parse_pages(
        self,
        pages: list,
        mark_lengths: dict[tuple[int, str], float] | None = None,
        page_default_lengths: dict[int, float] | None = None,
    ) -> pd.DataFrame:
        items: list[BOMItem] = []
        mark_lengths = mark_lengths or {}
        page_default_lengths = page_default_lengths or {}

        for page in pages:
            items.extend(self._parse_bom_tables(page.tables, page.page_number))
            items.extend(
                self._parse_drawing_marks(
                    page.text,
                    page.page_number,
                    mark_lengths,
                    page_default_lengths.get(page.page_number),
                )
            )
            items.extend(self._parse_steel_shapes(page.text, page.page_number))
            items.extend(
                self._parse_stacked_marks(
                    getattr(page, "stacked_marks", None),
                    page.page_number,
                    mark_lengths,
                    page_default_lengths.get(page.page_number),
                )
            )

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

    def _parse_drawing_marks(
        self,
        text: str,
        page_number: int,
        mark_lengths: dict[tuple[int, str], float] | None = None,
        page_default_length: float | None = None,
    ) -> list[BOMItem]:
        if not text.strip():
            return []

        mark_lengths = mark_lengths or {}

        # Count marks per line so a length/grade annotated on the same line as a
        # mark (e.g. "B3 ISMB250 LENGTH 3200") is captured, mirroring the steel
        # shape parser. Marks sharing the same length/grade are grouped together.
        counts: Counter[tuple[str, float | None, str]] = Counter()
        for line in text.splitlines():
            marks = DRAWING_MARK_PATTERN.findall(line)
            if not marks:
                continue
            length_match = LENGTH_PATTERN.search(line)
            length = float(length_match.group(1)) if length_match else None
            grade_match = GRADE_PATTERN.search(line)
            grade = grade_match.group(1).upper() if grade_match else ""
            for match in marks:
                counts[(self._normalize_mark(match), length, grade)] += 1

        items: list[BOMItem] = []
        for (mark, length, grade), count in sorted(
            counts.items(), key=lambda kv: (kv[0][0], kv[0][1] or 0.0, kv[0][2])
        ):
            # When the text layer has no length, fall back to a dimension read
            # from the drawing via OCR (per-mark first, then the page default
            # beam dimension for linear members).
            if length is None:
                length = mark_lengths.get((page_number, mark))
            if length is None and page_default_length is not None and _is_linear_mark(mark):
                length = page_default_length
            items.append(
                BOMItem(
                    mark=mark,
                    description=self._describe_mark(mark),
                    quantity=count,
                    length=length,
                    grade=grade,
                    source_page=page_number,
                )
            )
        return items

    def _parse_stacked_marks(
        self,
        stacked_marks: dict[str, int] | None,
        page_number: int,
        mark_lengths: dict[tuple[int, str], float] | None = None,
        page_default_length: float | None = None,
    ) -> list[BOMItem]:
        """Build items for marks recovered from rotated/stacked characters."""
        if not stacked_marks:
            return []
        mark_lengths = mark_lengths or {}

        items: list[BOMItem] = []
        for raw_mark, count in sorted(stacked_marks.items()):
            mark = self._normalize_mark(raw_mark)
            length = mark_lengths.get((page_number, mark))
            if length is None and page_default_length is not None and _is_linear_mark(mark):
                length = page_default_length
            items.append(
                BOMItem(
                    mark=mark,
                    description=self._describe_mark(mark),
                    quantity=count,
                    length=length,
                    grade="",
                    source_page=page_number,
                )
            )
        return items

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
