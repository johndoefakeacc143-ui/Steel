"""Extract structured steel members from raw drawing content."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from src.readers import RawContent
from src.sections import compute_weight_kg, normalize_section_type

logger = logging.getLogger(__name__)

NA = "NA"

# --- Table header synonyms -------------------------------------------------
HEADER_KEYWORDS: dict[str, tuple[str, ...]] = {
    "member_id": ("member id", "member_id", "member", "mark", "tag", "id", "ref", "item"),
    "member_type": ("member type", "type", "category", "element"),
    "section_size": ("section size", "section", "size", "profile", "designation"),
    "length_mm": ("length", "len", "l (mm)", "length (mm)", "length(m)", "length (m)", "l"),
    "quantity": ("quantity", "qty", "nos", "no.", "no", "count", "pcs", "nos."),
    "elevation_m": ("elevation", "elev", "el", "level", "rl", "e.l."),
    "grid_location": ("grid location", "grid", "gridline", "grid ref", "location", "axis"),
    "weight_kg": ("weight", "wt", "mass", "unit weight", "kg"),
}

# --- Member-type inference -------------------------------------------------
TYPE_PREFIXES: tuple[tuple[str, str], ...] = (
    ("BR", "Bracing"),
    ("BRC", "Bracing"),
    ("COL", "Column"),
    ("GDR", "Girder"),
    ("PL", "Plate"),
    ("G", "Girder"),
    ("C", "Column"),
    ("B", "Beam"),
    ("P", "Plate"),
)

SECTION_TEXT_RE = re.compile(
    r"\b((?:ISMB|ISMC|MB|MC|RHS|SHS|PL|PLATE|PLT)\s*[.:]?\s*"
    r"\d+(?:\.\d+)?(?:\s*[xX*×]\s*\d+(?:\.\d+)?){0,2})",
    re.IGNORECASE,
)
MEMBER_ID_RE = re.compile(r"\b([A-Z]{1,3}\d+[A-Za-z]?)\b")
LENGTH_RE = re.compile(
    r"(?:length|len|lg|l)\s*[:=]?\s*(\d+(?:\.\d+)?)\s*(mm|m|meter|metre)?",
    re.IGNORECASE,
)
QTY_RE = re.compile(r"(?:qty|quantity|nos|no\.?|count|pcs)\s*[:=]?\s*(\d+)", re.IGNORECASE)
ELEV_RE = re.compile(
    r"(?:elevation|elev|el|level|rl)\s*[:=]?\s*\(?\+?\)?\s*(\d+(?:\.\d+)?)\s*(mm|m)?",
    re.IGNORECASE,
)
GRID_RE = re.compile(r"(?:grid|axis|@)\s*[:=]?\s*([A-Z0-9]{1,3}\s*[-/]\s*[A-Z0-9]{1,3})", re.IGNORECASE)
WEIGHT_RE = re.compile(r"(?:weight|wt|mass)\s*[:=]?\s*(\d+(?:\.\d+)?)\s*(kg)?", re.IGNORECASE)
NUM_UNIT_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s*(mm|m|meter|metre|kg)?", re.IGNORECASE)


@dataclass
class Member:
    member_id: str | None = None
    member_type: str | None = None
    section_size: str | None = None
    length_mm: float | None = None
    quantity: int | None = None
    elevation_m: float | None = None
    grid_location: str | None = None
    weight_kg: float | None = None
    source_file: str = ""


class MemberExtractor:
    """Parse members from tables and free text within a ``RawContent``."""

    def extract(self, raw: RawContent) -> list[Member]:
        members: list[Member] = []
        for table in raw.tables:
            members.extend(self._from_table(table, raw.source))
        members.extend(self._from_text(raw.text, raw.source))
        return [self._finalize(m) for m in members if self._is_valid(m)]

    # -- Tables -------------------------------------------------------------
    def _from_table(self, table: list[list[str | None]], source: str) -> list[Member]:
        if not table or len(table) < 2:
            return []
        header_map = self._map_headers(table[0])
        if "member_id" not in header_map and "section_size" not in header_map:
            return []
        if len(header_map) < 2:
            return []

        members: list[Member] = []
        length_is_meter = self._length_header_in_meters(table[0], header_map)
        for row in table[1:]:
            if not row or all(not str(c).strip() for c in row if c is not None):
                continue
            members.append(self._row_to_member(row, header_map, source, length_is_meter))
        return members

    def _map_headers(self, header_row: list[str | None]) -> dict[str, int]:
        mapping: dict[str, int] = {}
        for index, cell in enumerate(header_row):
            if not cell:
                continue
            normalized = re.sub(r"\s+", " ", str(cell).strip().lower())
            for field, keywords in HEADER_KEYWORDS.items():
                if field in mapping:
                    continue
                if any(normalized == kw or normalized.startswith(kw + " ") or normalized == kw for kw in keywords) or any(
                    re.fullmatch(rf"{re.escape(kw)}s?", normalized) for kw in keywords
                ):
                    mapping[field] = index
                    break
        return mapping

    @staticmethod
    def _length_header_in_meters(header_row: list[str | None], header_map: dict[str, int]) -> bool:
        idx = header_map.get("length_mm")
        if idx is None or idx >= len(header_row):
            return False
        header = str(header_row[idx] or "").lower()
        return bool(re.search(r"\(m\)|\bm\b|meter|metre", header)) and "mm" not in header

    def _row_to_member(
        self,
        row: list[str | None],
        header_map: dict[str, int],
        source: str,
        length_is_meter: bool,
    ) -> Member:
        def cell(field: str) -> str:
            idx = header_map.get(field)
            if idx is None or idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()

        return Member(
            member_id=cell("member_id") or None,
            member_type=cell("member_type") or None,
            section_size=cell("section_size") or None,
            length_mm=self._parse_length(cell("length_mm"), force_meter=length_is_meter),
            quantity=self._parse_int(cell("quantity")),
            elevation_m=self._parse_elevation(cell("elevation_m")),
            grid_location=cell("grid_location") or None,
            weight_kg=self._parse_float(cell("weight_kg")),
            source_file=source,
        )

    # -- Text ---------------------------------------------------------------
    def _from_text(self, text: str, source: str) -> list[Member]:
        members: list[Member] = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            section_match = SECTION_TEXT_RE.search(line)
            if not section_match:
                continue
            section = section_match.group(1)
            remainder = line[: section_match.start()] + " " + line[section_match.end():]

            id_match = MEMBER_ID_RE.search(remainder)
            length_match = LENGTH_RE.search(line)
            qty_match = QTY_RE.search(line)
            elev_match = ELEV_RE.search(line)
            grid_match = GRID_RE.search(line)
            weight_match = WEIGHT_RE.search(line)

            members.append(
                Member(
                    member_id=id_match.group(1) if id_match else None,
                    section_size=re.sub(r"\s+", " ", section).strip(),
                    length_mm=self._parse_length_with_unit(length_match),
                    quantity=int(qty_match.group(1)) if qty_match else None,
                    elevation_m=self._parse_elev_match(elev_match),
                    grid_location=re.sub(r"\s+", "", grid_match.group(1)) if grid_match else None,
                    weight_kg=float(weight_match.group(1)) if weight_match else None,
                    source_file=source,
                )
            )
        return members

    # -- Finalisation -------------------------------------------------------
    def _finalize(self, member: Member) -> Member:
        if not member.member_type:
            member.member_type = self._infer_type(member)
        if member.quantity is None:
            member.quantity = 1
        if member.weight_kg is None and member.section_size:
            member.weight_kg = compute_weight_kg(member.section_size, member.length_mm)
        return member

    @staticmethod
    def _infer_type(member: Member) -> str | None:
        if member.section_size and normalize_section_type(member.section_size) == "PL":
            return "Plate"
        mid = (member.member_id or "").upper()
        for prefix, member_type in TYPE_PREFIXES:
            if mid.startswith(prefix):
                return member_type
        return None

    @staticmethod
    def _is_valid(member: Member) -> bool:
        return bool(member.member_id or member.section_size)

    # -- Value parsing ------------------------------------------------------
    def _parse_length(self, value: str, force_meter: bool) -> float | None:
        if not value:
            return None
        match = NUM_UNIT_RE.search(value)
        if not match:
            return None
        number = float(match.group(1))
        unit = (match.group(2) or "").lower()
        if unit == "mm":
            return number
        if unit in ("m", "meter", "metre") or force_meter:
            return number * 1000.0
        # No unit: values under 100 are implausible as mm for a member -> meters.
        return number * 1000.0 if number < 100 else number

    def _parse_length_with_unit(self, match: re.Match | None) -> float | None:
        if not match:
            return None
        number = float(match.group(1))
        unit = (match.group(2) or "").lower()
        if unit in ("m", "meter", "metre"):
            return number * 1000.0
        if unit == "mm":
            return number
        return number * 1000.0 if number < 100 else number

    def _parse_elevation(self, value: str) -> float | None:
        if not value:
            return None
        match = NUM_UNIT_RE.search(value)
        if not match:
            return None
        number = float(match.group(1))
        unit = (match.group(2) or "").lower()
        return number / 1000.0 if unit == "mm" else number

    @staticmethod
    def _parse_elev_match(match: re.Match | None) -> float | None:
        if not match:
            return None
        number = float(match.group(1))
        unit = (match.group(2) or "").lower()
        return number / 1000.0 if unit == "mm" else number

    @staticmethod
    def _parse_int(value: str) -> int | None:
        try:
            return int(float(re.sub(r"[^\d.]", "", value)))
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _parse_float(value: str) -> float | None:
        try:
            cleaned = re.sub(r"[^\d.]", "", value)
            return float(cleaned) if cleaned else None
        except (TypeError, ValueError):
            return None
