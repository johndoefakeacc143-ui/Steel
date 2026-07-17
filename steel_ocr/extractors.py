"""Extract beams, columns, and base plates from OCR / digital drawing text."""

from __future__ import annotations

import re
from typing import Any

from steel_ocr.patterns import (
    ANCHOR_BOLT_RE,
    BASEPLATE_MARK_RE,
    BEAM_MARK_RE,
    COLUMN_MARK_RE,
    ELEVATION_RE,
    LENGTH_RE,
    MATERIAL_RE,
    PLATE_SIZE_RE,
    SECTION_SIZE_RE,
    STANDALONE_LENGTH_RE,
)
from steel_ocr.reader import DocumentText, PageText


def normalize_section(raw: str) -> str:
    return re.sub(r"\s+", "", raw.upper().replace("×", "x"))


def find_first(pattern: re.Pattern[str], text: str, group: int = 1) -> str:
    m = pattern.search(text)
    if not m:
        return ""
    try:
        return (m.group(group) or "").strip()
    except IndexError:
        return m.group(0).strip()


def find_quantity(text: str, default: int = 1) -> int:
    """Parse Qty / Nos / (n) / trailing schedule count from a line."""
    # Prefer keyword forms (Qty=4)
    keyword = re.search(
        r"(?:Qty|QTY|Quantity|Count)\s*[:=]?\s*(\d{1,4})",
        text,
        re.IGNORECASE,
    )
    if keyword:
        value = int(keyword.group(1))
        if 1 <= value <= 500:
            return value

    nos = re.search(
        r"(\d{1,4})\s*(?:Nos?\.?|No\.?|Off|EA|Pcs?\.?)\b",
        text,
        re.IGNORECASE,
    )
    if nos:
        value = int(nos.group(1))
        if 1 <= value <= 500:
            return value

    # Parenthetical count — skip anchor-bolt style "(4) Ø22"
    for m in re.finditer(r"\(\s*(\d{1,4})\s*\)", text):
        after = text[m.end() : m.end() + 12]
        if re.match(r"\s*(?:M\d|Ø|DIA|mm|m)\b", after, re.IGNORECASE):
            continue
        value = int(m.group(1))
        if 1 <= value <= 500:
            return value

    # Tabular schedule: trailing small integer not part of a size/length/bolt token
    # e.g. "B1  W18x35  6000mm  4  A992" or "BP1  500x500x25  4-M20  8"
    cleaned = text
    cleaned = BEAM_MARK_RE.sub(" ", cleaned)
    cleaned = COLUMN_MARK_RE.sub(" ", cleaned)
    cleaned = BASEPLATE_MARK_RE.sub(" ", cleaned)
    cleaned = SECTION_SIZE_RE.sub(" ", cleaned)
    cleaned = PLATE_SIZE_RE.sub(" ", cleaned)
    cleaned = re.sub(r"\b\d{3,5}(?:\.\d+)?\s*mm\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b\d{1,3}'\s*-?\s*\d{1,2}\"?", " ", cleaned)
    cleaned = re.sub(
        r"\b\d+\s*[-–]\s*(?:M|Ø|DIA\.?\s*)\d+",
        " ",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"\bM\d{1,2}\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:EL|TOC|TOS)[^\s]*", " ", cleaned, flags=re.IGNORECASE)
    # Collect remaining small integers; prefer the last one on the line
    candidates = [
        int(m.group(0))
        for m in re.finditer(r"\b(\d{1,3})\b", cleaned)
        if 1 <= int(m.group(0)) <= 200
    ]
    if candidates:
        return candidates[-1]

    return default


def parse_anchor(text: str) -> tuple[str, str]:
    m = ANCHOR_BOLT_RE.search(text)
    if not m:
        return "", ""
    groups = m.groups()
    for i in range(0, len(groups), 2):
        qty, dia = groups[i], groups[i + 1] if i + 1 < len(groups) else None
        if qty and dia:
            return str(qty), str(dia)
    return "", ""


def mark_from_match(match: re.Match[str], kind: str) -> str:
    raw = re.sub(r"\s+", "", match.group(0)).upper()

    if kind == "beam":
        raw = re.sub(r"^BEAM-?", "", raw)
        if raw.startswith("BM"):
            return raw
        raw = re.sub(r"^B+", "B", raw)
        if not raw.startswith("B"):
            raw = "B" + raw
        return raw

    if kind == "column":
        raw = re.sub(r"^COLUMN-?", "", raw)
        if raw.startswith("COL"):
            suffix = re.sub(r"^COL-?", "", raw)
            return f"COL-{suffix}" if suffix else "COL"
        raw = re.sub(r"^C+", "C", raw)
        if not raw.startswith("C"):
            raw = "C" + raw
        return raw

    # base plate
    raw = raw.replace("BASEPLATE", "BP")
    raw = re.sub(r"^BASEPLATE", "BP", raw)
    if not raw.startswith("BP") and not raw.startswith("BPL"):
        raw = "BP" + re.sub(r"^BP?", "", raw)
    return raw


def iter_context_lines(text: str) -> list[str]:
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", text) if ln.strip()]
    expanded: list[str] = []
    for ln in lines:
        if len(ln) > 220 and "  " in ln:
            expanded.extend([p.strip() for p in re.split(r"\s{2,}", ln) if p.strip()])
        else:
            expanded.append(ln)
    return expanded


def line_section(line: str) -> str:
    section_m = SECTION_SIZE_RE.search(line)
    return normalize_section(section_m.group(0)) if section_m else ""


def extract_beams(text: str, page_num: int) -> list[dict[str, Any]]:
    beams: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in iter_context_lines(text):
        if BASEPLATE_MARK_RE.search(line) and not BEAM_MARK_RE.search(line):
            continue
        if re.search(r"\bCOLUMN\b", line, re.IGNORECASE) and not BEAM_MARK_RE.search(line):
            continue

        for mark_m in BEAM_MARK_RE.finditer(line):
            mark = mark_from_match(mark_m, "beam")
            if mark.startswith(("BP", "BR", "XB")) or mark in seen:
                continue
            if re.search(r"\b(?:BRACING|BRACE)\b", line, re.IGNORECASE):
                continue
            seen.add(mark)

            section = line_section(line)
            length = find_first(LENGTH_RE, line) or find_first(STANDALONE_LENGTH_RE, line)
            # LENGTH_RE also matches Qty — prefer dedicated quantity parser
            qty = find_quantity(line, default=1)
            # If length accidentally captured a bare qty digit, clear it
            if length.isdigit() and int(length) == qty and qty <= 50:
                length = find_first(STANDALONE_LENGTH_RE, line)

            beams.append(
                {
                    "Member Name": mark,
                    "Size": section,
                    "Quantity": qty,
                    "Length": length,
                    "Material": find_first(MATERIAL_RE, line, group=0),
                    "Page": page_num,
                    "Source Line": line[:160],
                }
            )

    # Plan sheets often label beams only by section (no B# mark)
    if not beams:
        beams.extend(_beams_from_sections_only(text, page_num))

    return beams


def _beams_from_sections_only(text: str, page_num: int) -> list[dict[str, Any]]:
    """Count section occurrences as anonymous beams when no marks found."""
    beams: list[dict[str, Any]] = []
    counts: dict[str, int] = {}
    for line in iter_context_lines(text):
        if BASEPLATE_MARK_RE.search(line) or COLUMN_MARK_RE.search(line):
            continue
        if re.search(r"\b(?:COLUMN|BASE\s*PLATE|BRACING)\b", line, re.IGNORECASE):
            continue
        for m in SECTION_SIZE_RE.finditer(line):
            section = normalize_section(m.group(0))
            counts[section] = counts.get(section, 0) + 1
    for idx, (section, qty) in enumerate(sorted(counts.items()), start=1):
        beams.append(
            {
                "Member Name": f"B-{section}-{idx}",
                "Size": section,
                "Quantity": qty,
                "Length": "",
                "Material": "",
                "Page": page_num,
                "Source Line": f"section label ×{qty}",
            }
        )
    return beams


def extract_columns(text: str, page_num: int) -> list[dict[str, Any]]:
    columns: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in iter_context_lines(text):
        if BASEPLATE_MARK_RE.search(line) and not COLUMN_MARK_RE.search(line):
            continue
        if re.search(r"\bBEAM\b", line, re.IGNORECASE) and not re.search(
            r"\b(?:COLUMN|COL|C)\d", line, re.IGNORECASE
        ):
            continue

        for mark_m in COLUMN_MARK_RE.finditer(line):
            mark = mark_from_match(mark_m, "column")
            if mark in seen:
                continue
            section = line_section(line)
            # Require section nearby OR explicit COLUMN/COL word to avoid random C tokens
            if not section and not re.search(r"\b(?:COLUMN|COL)\b", line, re.IGNORECASE):
                # Still accept C1-style marks with digits when line has steel context
                if not re.search(r"\bC\d", mark, re.IGNORECASE):
                    continue
                if not section and len(line) < 8:
                    continue
            seen.add(mark)

            height = find_first(LENGTH_RE, line) or find_first(STANDALONE_LENGTH_RE, line)
            qty = find_quantity(line, default=1)
            if height.isdigit() and int(height) == qty and qty <= 50:
                height = find_first(STANDALONE_LENGTH_RE, line)

            columns.append(
                {
                    "Member Name": mark,
                    "Size": section,
                    "Quantity": qty,
                    "Height": height,
                    "Material": find_first(MATERIAL_RE, line, group=0),
                    "Page": page_num,
                    "Source Line": line[:160],
                }
            )
    return columns


def extract_base_plates(text: str, page_num: int) -> list[dict[str, Any]]:
    plates: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in iter_context_lines(text):
        for mark_m in BASEPLATE_MARK_RE.finditer(line):
            mark = mark_from_match(mark_m, "baseplate")
            if mark in seen:
                continue
            seen.add(mark)

            plate_m = PLATE_SIZE_RE.search(line)
            plate_size = ""
            thickness = ""
            if plate_m:
                plate_size = plate_m.group(1).replace("×", "x").replace(" ", "")
                parts = re.split(r"[xX]", plate_size)
                if len(parts) >= 3:
                    thickness = parts[-1]
            if not thickness:
                th = re.search(r"Thickness\s*[:=]?\s*(\d+(?:\.\d+)?)", line, re.IGNORECASE)
                if th:
                    thickness = th.group(1)

            qty_bolts, dia = parse_anchor(line)
            elevs = [m.group(1).strip() for m in ELEVATION_RE.finditer(line)]
            member_qty = find_quantity(line, default=1)

            size_display = plate_size
            if thickness and thickness not in size_display:
                size_display = plate_size or f"t={thickness}"

            plates.append(
                {
                    "Member Name": mark,
                    "Size": size_display,
                    "Quantity": member_qty,
                    "Thickness": thickness,
                    "Anchor Bolt Dia": dia,
                    "Anchor Bolt Qty": qty_bolts,
                    "Top of Concrete EL": elevs[0] if elevs else "",
                    "Page": page_num,
                    "Source Line": line[:160],
                }
            )
    return plates


def extract_page_members(page: PageText) -> dict[str, list[dict[str, Any]]]:
    text = page.text or ""
    page_num = page.page_number
    return {
        "beams": extract_beams(text, page_num),
        "columns": extract_columns(text, page_num),
        "base_plates": extract_base_plates(text, page_num),
    }


def extract_document_members(doc: DocumentText) -> dict[str, list[dict[str, Any]]]:
    """Trace every page and collect all beams, columns, and base plates."""
    beams: list[dict[str, Any]] = []
    columns: list[dict[str, Any]] = []
    plates: list[dict[str, Any]] = []

    for page in doc.pages:
        result = extract_page_members(page)
        beams.extend(result["beams"])
        columns.extend(result["columns"])
        plates.extend(result["base_plates"])

    return {
        "beams": beams,
        "columns": columns,
        "base_plates": plates,
    }
