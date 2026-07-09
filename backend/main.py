"""
SteelDraw AI Extractor — FastAPI backend

Upload large steel structure engineering PDFs, extract Beams / Columns / Base Plates
page-by-page, and return a multi-sheet Excel workbook.
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import traceback
from pathlib import Path
from typing import Any

import cv2
import numpy as np
import pandas as pd
import pdfplumber
import pytesseract
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from PIL import Image

# ---------------------------------------------------------------------------
# Env / app setup
# ---------------------------------------------------------------------------

load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")
load_dotenv()  # also allow backend/.env

MAX_UPLOAD_BYTES = 500 * 1024 * 1024  # 500 MB
DIGITAL_TEXT_THRESHOLD = 40  # chars/page below this → treat as scanned

app = FastAPI(
    title="SteelDraw AI Extractor",
    description="Extract beams, columns, and base plates from steel drawings",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ===========================================================================
# REGEX PATTERNS — edit these to match your drawing standards
# ===========================================================================

# Section sizes: W18x35, W21×44, ISMB400, ISMC300, UB457x191x67, UC254x254x73, HSS6x6x3/8, etc.
SECTION_SIZE_RE = re.compile(
    r"\b(?:"
    r"W\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # AISC W shapes
    r"|HP\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # HP piles
    r"|WT\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # WT tees
    r"|C\d{1,2}\s*[x×]\s*\d{1,2}(?:\.\d+)?"  # Channels
    r"|MC\d{1,2}\s*[x×]\s*\d{1,2}(?:\.\d+)?"
    r"|L\d{1,2}\s*[x×]\s*\d{1,2}\s*[x×]\s*\d/\d+"  # Angles
    r"|HSS\d{1,2}(?:\.\d+)?\s*[x×]\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d/]+"  # HSS
    r"|ISMB\s?\d{2,4}"  # Indian Standard Medium Beam
    r"|ISMC\s?\d{2,4}"  # Indian Standard Medium Channel
    r"|ISWB\s?\d{2,4}"
    r"|UB\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"  # Universal Beam (UK)
    r"|UC\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"  # Universal Column
    r"|HEA\s?\d{2,3}"
    r"|HEB\s?\d{2,3}"
    r"|IPE\s?\d{2,3}"
    r")\b",
    re.IGNORECASE,
)

# Piece marks: B1, B-12, BM-3A, C1, COL-2, BP1, BP-01, etc.
# NOTE: Beam pattern excludes BP (base plates) and BR/XB (bracing).
BEAM_MARK_RE = re.compile(
    r"\b(?:BEAM|BM|B(?![PR]))([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)
COLUMN_MARK_RE = re.compile(
    r"\b(?:COLUMN|COL|C)([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)
BASEPLATE_MARK_RE = re.compile(
    r"\b(?:BPL|BP|BASE\s*PLATE)([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)
# Bracing marks: BR1, BR-2, BRACE-3, XB1, etc.
BRACING_MARK_RE = re.compile(
    r"\b(?:BRACE|BRACING|BR|XB)([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

# Page-type keywords — edit if your title blocks use different wording
PLAN_PAGE_RE = re.compile(
    r"\b(?:FLOOR\s+FRAMING|FRAMING\s+PLAN|ROOF\s+PLAN|FOUNDATION\s+PLAN|"
    r"STRUCTURAL\s+PLAN|PLAN\s*[-–]|PLAN\b|"
    r"UTILITY\s+SUPPORT\s+PLAN)\b",
    re.IGNORECASE,
)
ELEVATION_PAGE_RE = re.compile(
    r"\b(?:ELEVATION|BRACE\s+FRAME\s+ELEV|FRAME\s+ELEVATION|"
    r"BUILDING\s+ELEVATION|SECTION\s+ELEVATION)\b",
    re.IGNORECASE,
)

# Length / height: 12'-6", 3750mm, 12.5 ft, 4500, L=6000
LENGTH_RE = re.compile(
    r"(?:L(?:ength)?|Ht|H(?:eight)?)\s*[:=]?\s*"
    r"("
    r"\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?"  # imperial feet-inches
    r"|\d{3,5}(?:\.\d+)?\s*(?:mm|m)?"  # metric
    r"|\d{1,3}(?:\.\d+)?\s*(?:ft|FT)"
    r")",
    re.IGNORECASE,
)
STANDALONE_LENGTH_RE = re.compile(
    r"\b(\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?|\d{3,5}\s*mm)\b",
    re.IGNORECASE,
)

# Elevations: EL +12.500, EL. 100'-0", T.O.S. +25.000, TOC EL 0.000
# Prefer imperial feet-inches first so "0'-0\"" is not truncated to "0".
# Metric unit group uses \b so trailing "M" from "Material" is not captured.
ELEVATION_RE = re.compile(
    r"(?:EL(?:EV(?:ATION)?)?|T\.?O\.?S\.?|T\.?O\.?C\.?|TOP|BASE)\s*[.:]?\s*"
    r"([+-]?\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?"
    r"|[+-]?\d{1,4}(?:\.\d{1,3})?(?:\s*(?:mm|m|ft)\b)?)",
    re.IGNORECASE,
)

# Material grades
MATERIAL_RE = re.compile(
    r"\b(A36|A572(?:\s*Gr\.?\s*\d+)?|A992|A500(?:\s*Gr\.?\s*[ABC])?|"
    r"S275(?:JR)?|S355(?:JR)?|ASTM\s*A\d+|IS\s*2062(?:\s*E\d+)?|"
    r"Gr\.?\s*(?:36|50|55)|FY\s*\d{2,3})\b",
    re.IGNORECASE,
)

# Base plate size: 600x600x30, 24"x24"x1-1/4", PL 500x500x25
# Negative lookbehind avoids stealing UC254x254x73 / UB / HSS section sizes.
PLATE_SIZE_RE = re.compile(
    r"(?<![A-Z])(?:PL(?:ATE)?\s*)?"
    r"(\d{2,4}(?:\.\d+)?\s*[x×]\s*\d{2,4}(?:\.\d+)?\s*[x×]\s*\d{1,3}(?:\.\d+)?)"
    r"(?:\s*mm)?",
    re.IGNORECASE,
)

# Anchor bolts: 4-M20, (4) Ø22, 4 Nos M24, 6-3/4" dia
ANCHOR_BOLT_RE = re.compile(
    r"(?:"
    r"(\d+)\s*[-–]\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"  # 4-M20 / 6-M16
    r"|(\d+)\s*(?:Nos?\.?|No\.?)\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"  # 4 Nos M24
    r"|\(\s*(\d+)\s*\)\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"  # (4) Ø22
    r"|(\d+)\s*[-–]\s*(\d/\d|\d(?:\.\d+)?)\s*[\"']?\s*(?:DIA|Ø)?"  # 6-3/4" dia
    r")",
    re.IGNORECASE,
)


# ===========================================================================
# PDF type detection + OCR
# ===========================================================================

def page_has_digital_text(page: pdfplumber.page.Page) -> bool:
    """Return True if the page looks like a digital (text-based) PDF."""
    try:
        text = page.extract_text() or ""
        return len(text.strip()) >= DIGITAL_TEXT_THRESHOLD
    except Exception:
        return False


def preprocess_for_ocr(bgr: np.ndarray) -> np.ndarray:
    """OpenCV preprocessing to improve OCR on scanned drawings."""
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    # Upscale thin linework
    gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    # Denoise + adaptive threshold for blueprint-style scans
    denoise = cv2.fastNlMeansDenoising(gray, h=10)
    binary = cv2.adaptiveThreshold(
        denoise,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        11,
    )
    return binary


def ocr_page_image(page: pdfplumber.page.Page, resolution: int = 200) -> str:
    """Render a PDF page to image, preprocess, and OCR with Tesseract."""
    try:
        pil_img: Image.Image = page.to_image(resolution=resolution).original
        bgr = cv2.cvtColor(np.array(pil_img.convert("RGB")), cv2.COLOR_RGB2BGR)
        processed = preprocess_for_ocr(bgr)
        config = "--oem 3 --psm 6"
        text = pytesseract.image_to_string(processed, config=config)
        return text or ""
    except Exception as exc:
        print(f"[OCR] page failed: {exc}")
        return ""


def extract_page_text(page: pdfplumber.page.Page) -> tuple[str, str]:
    """
    Extract text from a page.
    Returns (text, source) where source is 'digital' or 'ocr'.
    """
    if page_has_digital_text(page):
        return (page.extract_text() or "", "digital")
    return (ocr_page_image(page), "ocr")


def detect_page_type(text: str) -> str:
    """
    Classify a drawing page as Plan, Elevation, or Other.

    Priority: Elevation title wins over Plan if both appear (e.g. keyplan
    notes on an elevation sheet). Edit PLAN_PAGE_RE / ELEVATION_PAGE_RE above.
    """
    head = (text or "")[:2500]
    elev_hit = bool(ELEVATION_PAGE_RE.search(head))
    plan_hit = bool(PLAN_PAGE_RE.search(head))
    if elev_hit and not plan_hit:
        return "Elevation"
    if plan_hit and not elev_hit:
        return "Plan"
    if elev_hit and plan_hit:
        # Prefer whichever keyword appears first in the title block area
        elev_pos = ELEVATION_PAGE_RE.search(head)
        plan_pos = PLAN_PAGE_RE.search(head)
        if elev_pos and plan_pos:
            return "Elevation" if elev_pos.start() < plan_pos.start() else "Plan"
        return "Elevation" if elev_pos else "Plan"
    # Heuristic fallback from content density
    beamish = len(SECTION_SIZE_RE.findall(text or ""))
    colish = len(COLUMN_MARK_RE.findall(text or ""))
    if colish >= 2 and beamish <= colish:
        return "Elevation"
    if beamish >= 3:
        return "Plan"
    return "Other"


def _format_length_ft_in(meters: float) -> str:
    """Format meters as feet-inches for display (e.g. 45.72 -> 150'-0\")."""
    total_inches = meters / 0.0254
    feet = int(total_inches // 12)
    inches = int(round(total_inches % 12))
    if inches == 12:
        feet += 1
        inches = 0
    return f"{feet}'-{inches}\""


def _total_length_stats(rows: list[dict[str, Any]], length_key: str = "Length") -> dict[str, Any]:
    total_m = 0.0
    counted = 0
    for row in rows:
        length_m = _parse_length_to_meters(str(row.get(length_key, "")))
        if length_m is not None:
            total_m += length_m
            counted += 1
    return {
        "count": len(rows),
        "with_length": counted,
        "total_length_m": round(total_m, 3) if total_m else 0.0,
        "total_length_ft_in": _format_length_ft_in(total_m) if total_m else "N/A",
    }


# ===========================================================================
# Camelot table extraction (optional — fails gracefully if ghostscript missing)
# ===========================================================================

def extract_tables_text(pdf_path: str, page_number: int) -> str:
    """
    Try Camelot lattice/stream tables; return concatenated cell text.

    Camelot is optional. It is not in requirements.txt because camelot-py[cv]
    depends on pdftopng, which often fails to install on Windows.
    """
    try:
        import camelot  # type: ignore  # optional dependency
    except ImportError:
        return ""

    try:
        tables = camelot.read_pdf(
            pdf_path,
            pages=str(page_number),
            flavor="lattice",
            suppress_stdout=True,
        )
        if tables.n == 0:
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_number),
                flavor="stream",
                suppress_stdout=True,
            )
        chunks: list[str] = []
        for table in tables:
            try:
                chunks.append(table.df.to_string(index=False, header=False))
            except Exception:
                continue
        return "\n".join(chunks)
    except Exception as exc:
        print(f"[Camelot] page {page_number}: {exc}")
        return ""


# ===========================================================================
# Regex extractors
# ===========================================================================

def _normalize_section(raw: str) -> str:
    return re.sub(r"\s+", "", raw.upper().replace("×", "x"))


def _find_first(pattern: re.Pattern[str], text: str, group: int = 1) -> str:
    m = pattern.search(text)
    if not m:
        return ""
    try:
        return (m.group(group) or "").strip()
    except IndexError:
        return m.group(0).strip()


def _find_all_elevations(text: str) -> list[str]:
    return [m.group(1).strip() for m in ELEVATION_RE.finditer(text)]


def _parse_anchor(text: str) -> tuple[str, str]:
    m = ANCHOR_BOLT_RE.search(text)
    if not m:
        return "", ""
    groups = m.groups()
    # Patterns return (qty, dia) in alternating pairs
    for i in range(0, len(groups), 2):
        qty, dia = groups[i], groups[i + 1] if i + 1 < len(groups) else None
        if qty and dia:
            return str(qty), str(dia)
    return "", ""


def _mark_from_match(match: re.Match[str], kind: str) -> str:
    """
    Build a clean mark from a regex match.

    Handles compact marks (B4, C1) and verbose labels ("Beam B4", "Column C4")
    without doubling the letter prefix.
    """
    raw = re.sub(r"\s+", "", match.group(0)).upper()

    if kind == "beam":
        # BEAMB4 / BEAM-B4 / BEAM4 -> B4 ; BM-3A stays ; B12 stays
        raw = re.sub(r"^BEAM-?", "", raw)
        if raw.startswith("BM"):
            return raw
        raw = re.sub(r"^B+", "B", raw)  # BB4 -> B4
        if not raw.startswith("B"):
            raw = "B" + raw
        return raw

    if kind == "column":
        raw = re.sub(r"^COLUMN-?", "", raw)
        if raw.startswith("COL"):
            suffix = re.sub(r"^COL-?", "", raw)
            return f"COL-{suffix}" if suffix else "COL"
        raw = re.sub(r"^C+", "C", raw)  # CC4 -> C4
        if not raw.startswith("C"):
            raw = "C" + raw
        return raw

    if kind == "bracing":
        raw = re.sub(r"^BRACING-?", "BR", raw)
        raw = re.sub(r"^BRACE-?", "BR", raw)
        if raw.startswith("XB"):
            return raw
        raw = re.sub(r"^BR+", "BR", raw)
        if not raw.startswith("BR"):
            raw = "BR" + raw
        return raw

    # base plate
    raw = raw.replace("BASEPLATE", "BP")
    return raw


def _iter_context_lines(text: str) -> list[str]:
    """
    Split drawing text into logical lines for field association.

    Schedule rows are usually one member per line; using the whole page as a
    window causes neighboring rows to bleed into each other.
    """
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", text) if ln.strip()]
    # Also split very long OCR blobs on multiple spaces / tabs
    expanded: list[str] = []
    for ln in lines:
        if len(ln) > 220 and "  " in ln:
            expanded.extend([p.strip() for p in re.split(r"\s{2,}", ln) if p.strip()])
        else:
            expanded.append(ln)
    return expanded


def _line_section(line: str) -> str:
    section_m = SECTION_SIZE_RE.search(line)
    return _normalize_section(section_m.group(0)) if section_m else ""


def extract_beams_regex(text: str, page_num: int) -> list[dict[str, Any]]:
    beams: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in _iter_context_lines(text):
        # Skip obvious base-plate / column schedule rows
        if BASEPLATE_MARK_RE.search(line) and not BEAM_MARK_RE.search(line):
            continue
        if re.search(r"\bCOLUMN\b", line, re.IGNORECASE) and not BEAM_MARK_RE.search(line):
            continue

        for mark_m in BEAM_MARK_RE.finditer(line):
            mark = _mark_from_match(mark_m, "beam")
            if mark.startswith(("BP", "BR", "XB")) or mark in seen:
                continue
            # Skip lines that are clearly bracing rows
            if re.search(r"\b(?:BRACING|BRACE)\b", line, re.IGNORECASE):
                continue
            seen.add(mark)

            section = _line_section(line)
            length = _find_first(LENGTH_RE, line) or _find_first(STANDALONE_LENGTH_RE, line)
            material = _find_first(MATERIAL_RE, line, group=0)
            elevs = _find_all_elevations(line)
            start_el = elevs[0] if elevs else ""
            end_el = elevs[1] if len(elevs) > 1 else (elevs[0] if elevs else "")

            beams.append(
                {
                    "Mark": mark,
                    "Section Size": section,
                    "Length": length,
                    "Material": material,
                    "Start EL": start_el,
                    "End EL": end_el,
                    "Page": page_num,
                }
            )
    return beams


def extract_columns_regex(text: str, page_num: int) -> list[dict[str, Any]]:
    columns: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in _iter_context_lines(text):
        if BASEPLATE_MARK_RE.search(line) and not COLUMN_MARK_RE.search(line):
            continue
        # Avoid beam schedule lines that only mention "C" inside other words
        if re.search(r"\bBEAM\b", line, re.IGNORECASE) and not re.search(
            r"\b(?:COLUMN|COL|C)\d", line, re.IGNORECASE
        ):
            continue

        for mark_m in COLUMN_MARK_RE.finditer(line):
            mark = _mark_from_match(mark_m, "column")
            if mark in seen:
                continue
            # Require a steel section nearby so random "C" tokens are ignored
            section = _line_section(line)
            if not section and not re.search(r"\b(?:COLUMN|COL)\b", line, re.IGNORECASE):
                continue
            seen.add(mark)

            height = _find_first(LENGTH_RE, line) or _find_first(STANDALONE_LENGTH_RE, line)
            material = _find_first(MATERIAL_RE, line, group=0)
            elevs = _find_all_elevations(line)
            base_el = elevs[0] if elevs else ""
            top_el = elevs[1] if len(elevs) > 1 else ""

            columns.append(
                {
                    "Mark": mark,
                    "Section Size": section,
                    "Height": height,
                    "Base Elevation": base_el,
                    "Top Elevation": top_el,
                    "Material": material,
                    "Page": page_num,
                }
            )
    return columns


def extract_baseplates_regex(text: str, page_num: int) -> list[dict[str, Any]]:
    plates: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in _iter_context_lines(text):
        for mark_m in BASEPLATE_MARK_RE.finditer(line):
            mark = _mark_from_match(mark_m, "baseplate")
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
            # Explicit "Thickness 30" fallback
            if not thickness:
                th = re.search(r"Thickness\s*[:=]?\s*(\d+(?:\.\d+)?)", line, re.IGNORECASE)
                if th:
                    thickness = th.group(1)

            qty, dia = _parse_anchor(line)
            elevs = _find_all_elevations(line)
            toc_el = elevs[0] if elevs else ""

            plates.append(
                {
                    "Mark": mark,
                    "Plate Size": plate_size,
                    "Thickness": thickness,
                    "Anchor Bolt Dia": dia,
                    "Anchor Bolt Qty": qty,
                    "Top of Concrete EL": toc_el,
                    "Page": page_num,
                }
            )
    return plates


def extract_bracing_regex(text: str, page_num: int) -> list[dict[str, Any]]:
    """Extract bracing members (BR1, BRACE-2, HSS braces on plan)."""
    braces: list[dict[str, Any]] = []
    seen: set[str] = set()

    for line in _iter_context_lines(text):
        # Skip legend-only lines that mention bracing without a member
        if re.search(r"\bBRACE\s+FRAME\s+PER\b", line, re.IGNORECASE):
            continue

        mark_matches = list(BRACING_MARK_RE.finditer(line))
        if mark_matches:
            for mark_m in mark_matches:
                mark = _mark_from_match(mark_m, "bracing")
                # Avoid capturing bare "BRACING" / "BRACE" without an id
                if not re.search(r"\d", mark):
                    continue
                if mark in seen:
                    continue
                seen.add(mark)
                section = _line_section(line)
                length = _find_first(LENGTH_RE, line) or _find_first(STANDALONE_LENGTH_RE, line)
                material = _find_first(MATERIAL_RE, line, group=0)
                braces.append(
                    {
                        "Mark": mark,
                        "Section Size": section,
                        "Length": length,
                        "Material": material,
                        "Page": page_num,
                    }
                )
            continue

        # Framing plans often label braces only by section + "bracing" word
        if re.search(r"\b(?:BRACING|BRACE|DIAGONAL)\b", line, re.IGNORECASE):
            section = _line_section(line)
            if not section:
                continue
            length = _find_first(LENGTH_RE, line) or _find_first(STANDALONE_LENGTH_RE, line)
            mark = f"BR-{section}-{len(seen) + 1}"
            if mark in seen:
                continue
            seen.add(mark)
            braces.append(
                {
                    "Mark": mark,
                    "Section Size": section,
                    "Length": length,
                    "Material": _find_first(MATERIAL_RE, line, group=0),
                    "Page": page_num,
                }
            )
    return braces


def extract_plan_beams_from_sections(text: str, page_num: int) -> list[dict[str, Any]]:
    """
    On framing PLAN sheets, beams are often labeled only as W24x76 (no B1 mark).
    Count each section occurrence on its line as a beam instance.
    """
    beams: list[dict[str, Any]] = []
    # Skip lines that are clearly bracing / columns / base plates
    for idx, line in enumerate(_iter_context_lines(text), start=1):
        if BRACING_MARK_RE.search(line) or re.search(
            r"\b(?:BRACING|BRACE|COLUMN|BASE\s*PLATE)\b", line, re.IGNORECASE
        ):
            continue
        if COLUMN_MARK_RE.search(line) and not BEAM_MARK_RE.search(line):
            continue
        # Already captured via piece-mark extractor
        if BEAM_MARK_RE.search(line):
            continue

        for sec_m in SECTION_SIZE_RE.finditer(line):
            section = _normalize_section(sec_m.group(0))
            # Prefer W / UB / ISMB style as beams; skip pure HSS on plan unless no brace word
            if section.startswith("HSS") and re.search(r"\b(?:BRACE|BRACING)\b", line, re.IGNORECASE):
                continue
            length = _find_first(LENGTH_RE, line) or _find_first(STANDALONE_LENGTH_RE, line)
            material = _find_first(MATERIAL_RE, line, group=0)
            elevs = _find_all_elevations(line)
            mark = f"BM-{section}-P{page_num}-{idx}"
            beams.append(
                {
                    "Mark": mark,
                    "Section Size": section,
                    "Length": length,
                    "Material": material,
                    "Start EL": elevs[0] if elevs else "",
                    "End EL": elevs[1] if len(elevs) > 1 else (elevs[0] if elevs else ""),
                    "Page": page_num,
                }
            )
    return beams


# ===========================================================================
# AI extraction (LangChain + OpenAI) — supplements regex
# ===========================================================================

AI_SYSTEM_PROMPT = """You are an expert steel detailer / structural steel modeler.
Extract structural steel members from drawing text.
Return ONLY valid JSON with this exact shape:
{
  "beams": [{"mark":"", "section_size":"", "length":"", "material":"", "start_el":"", "end_el":""}],
  "columns": [{"mark":"", "section_size":"", "height":"", "base_elevation":"", "top_elevation":"", "material":""}],
  "base_plates": [{"mark":"", "plate_size":"", "thickness":"", "anchor_bolt_dia":"", "anchor_bolt_qty":"", "top_of_concrete_el":""}]
}
Use empty string for unknown fields. Do not invent members not supported by the text.
Marks look like B1, BM-12, C3, COL-2, BP1. Sections like W18x35, ISMB400. Plates like 600x600x30.
"""


def _openai_api_key() -> str:
    """Return a usable OpenAI key, ignoring placeholders from .env.example."""
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        return ""
    placeholders = {
        "sk-your-key-here",
        "sk-your-real-key",
        "your-api-key",
        "changeme",
    }
    if api_key.lower() in placeholders or api_key.endswith("your-key-here"):
        return ""
    return api_key


def ai_extract_from_text(text: str, page_num: int) -> dict[str, list[dict[str, Any]]]:
    """Call OpenAI via LangChain when OPENAI_API_KEY is set; otherwise skip."""
    api_key = _openai_api_key()
    if not api_key or not text.strip():
        return {"beams": [], "columns": [], "base_plates": []}

    # Truncate very large pages to stay within context limits
    snippet = text[:12000]

    try:
        from langchain_openai import ChatOpenAI
        from langchain_core.messages import HumanMessage, SystemMessage

        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        llm = ChatOpenAI(model=model_name, temperature=0, api_key=api_key)
        response = llm.invoke(
            [
                SystemMessage(content=AI_SYSTEM_PROMPT),
                HumanMessage(
                    content=f"Page {page_num} drawing text:\n\n{snippet}\n\nExtract members as JSON."
                ),
            ]
        )
        content = response.content if isinstance(response.content, str) else str(response.content)
        return _parse_ai_json(content, page_num)
    except Exception as exc:
        print(f"[AI] page {page_num}: {exc}")
        return {"beams": [], "columns": [], "base_plates": []}


def _parse_ai_json(content: str, page_num: int) -> dict[str, list[dict[str, Any]]]:
    import json

    # Strip markdown fences if present
    cleaned = content.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        # Try to find a JSON object in the response
        m = re.search(r"\{[\s\S]*\}", cleaned)
        if not m:
            return {"beams": [], "columns": [], "base_plates": []}
        try:
            data = json.loads(m.group(0))
        except json.JSONDecodeError:
            return {"beams": [], "columns": [], "base_plates": []}

    beams = []
    for b in data.get("beams") or []:
        if not isinstance(b, dict):
            continue
        mark = str(b.get("mark") or "").strip()
        if not mark:
            continue
        beams.append(
            {
                "Mark": mark.upper(),
                "Section Size": str(b.get("section_size") or "").strip(),
                "Length": str(b.get("length") or "").strip(),
                "Material": str(b.get("material") or "").strip(),
                "Start EL": str(b.get("start_el") or "").strip(),
                "End EL": str(b.get("end_el") or "").strip(),
                "Page": page_num,
            }
        )

    columns = []
    for c in data.get("columns") or []:
        if not isinstance(c, dict):
            continue
        mark = str(c.get("mark") or "").strip()
        if not mark:
            continue
        columns.append(
            {
                "Mark": mark.upper(),
                "Section Size": str(c.get("section_size") or "").strip(),
                "Height": str(c.get("height") or "").strip(),
                "Base Elevation": str(c.get("base_elevation") or "").strip(),
                "Top Elevation": str(c.get("top_elevation") or "").strip(),
                "Material": str(c.get("material") or "").strip(),
                "Page": page_num,
            }
        )

    plates = []
    for p in data.get("base_plates") or []:
        if not isinstance(p, dict):
            continue
        mark = str(p.get("mark") or "").strip()
        if not mark:
            continue
        plates.append(
            {
                "Mark": mark.upper(),
                "Plate Size": str(p.get("plate_size") or "").strip(),
                "Thickness": str(p.get("thickness") or "").strip(),
                "Anchor Bolt Dia": str(p.get("anchor_bolt_dia") or "").strip(),
                "Anchor Bolt Qty": str(p.get("anchor_bolt_qty") or "").strip(),
                "Top of Concrete EL": str(p.get("top_of_concrete_el") or "").strip(),
                "Page": page_num,
            }
        )

    return {"beams": beams, "columns": columns, "base_plates": plates}


# ===========================================================================
# Merge / dedupe helpers
# ===========================================================================

def _merge_by_mark(items: list[dict[str, Any]], mark_key: str = "Mark") -> list[dict[str, Any]]:
    """Prefer rows with more filled fields when marks collide."""
    best: dict[str, dict[str, Any]] = {}
    for item in items:
        mark = str(item.get(mark_key, "")).strip().upper()
        if not mark:
            continue
        item = {**item, mark_key: mark}
        filled = sum(1 for k, v in item.items() if k != "Page" and str(v).strip())
        prev = best.get(mark)
        if prev is None:
            best[mark] = item
            continue
        prev_filled = sum(1 for k, v in prev.items() if k != "Page" and str(v).strip())
        if filled > prev_filled:
            # Keep earlier page if new one doesn't add page info preference
            best[mark] = item
        else:
            # Fill empty fields from the new item
            merged = dict(prev)
            for k, v in item.items():
                if not str(merged.get(k, "")).strip() and str(v).strip():
                    merged[k] = v
            best[mark] = merged
    return list(best.values())


# ===========================================================================
# Excel export
# ===========================================================================

def _parse_numeric_elevation(value: str) -> float | None:
    if not value:
        return None
    s = str(value).strip()
    # metric like +12.500 or 12500 mm
    m = re.search(r"([+-]?\d+(?:\.\d+)?)", s.replace(",", ""))
    if not m:
        return None
    num = float(m.group(1))
    if "mm" in s.lower():
        return num / 1000.0
    return num


def _parse_length_to_meters(value: str) -> float | None:
    if not value:
        return None
    s = str(value).strip()
    # 12'-6"
    imperial = re.match(
        r"(\d{1,3})'\s*-?\s*(\d{1,2})(?:\s*(\d)/(\d))?\"?",
        s,
    )
    if imperial:
        feet = float(imperial.group(1))
        inches = float(imperial.group(2))
        if imperial.group(3) and imperial.group(4):
            inches += float(imperial.group(3)) / float(imperial.group(4))
        return (feet + inches / 12.0) * 0.3048
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mm|m|ft)?", s, re.IGNORECASE)
    if not m:
        return None
    num = float(m.group(1))
    unit = (m.group(2) or "").lower()
    if unit == "mm":
        return num / 1000.0
    if unit == "ft":
        return num * 0.3048
    if unit == "m":
        return num
    # bare large numbers often mm on steel drawings
    if num >= 100:
        return num / 1000.0
    return num


def build_view_metrics(
    page_types: list[dict[str, Any]],
    beams: list[dict[str, Any]],
    bracing: list[dict[str, Any]],
    columns: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    Build Plan vs Elevation metrics the UI shows first.

    Plan pages  -> beam count/length + bracing count/length
    Elevation   -> column count/length (height)
    """
    plan_pages = [p["page"] for p in page_types if p.get("type") == "Plan"]
    elev_pages = [p["page"] for p in page_types if p.get("type") == "Elevation"]

    plan_beams = [b for b in beams if b.get("Page") in plan_pages] if plan_pages else list(beams)
    # If we found plan pages, prefer plan-only beams; else keep all beams for totals
    if plan_pages:
        beams_for_plan = plan_beams
    else:
        beams_for_plan = beams

    if plan_pages:
        bracing_for_plan = [b for b in bracing if b.get("Page") in plan_pages]
    else:
        bracing_for_plan = bracing

    if elev_pages:
        columns_for_elev = [c for c in columns if c.get("Page") in elev_pages]
    else:
        columns_for_elev = columns

    beam_stats = _total_length_stats(beams_for_plan, "Length")
    brace_stats = _total_length_stats(bracing_for_plan, "Length")
    # Columns use Height as length
    col_stats = _total_length_stats(
        [{**c, "Length": c.get("Height", "")} for c in columns_for_elev],
        "Length",
    )

    return {
        "plan_pages": plan_pages,
        "elevation_pages": elev_pages,
        "plan": {
            "beam_count": beam_stats["count"],
            "beam_length_m": beam_stats["total_length_m"],
            "beam_length_ft_in": beam_stats["total_length_ft_in"],
            "bracing_count": brace_stats["count"],
            "bracing_length_m": brace_stats["total_length_m"],
            "bracing_length_ft_in": brace_stats["total_length_ft_in"],
        },
        "elevation": {
            "column_count": col_stats["count"],
            "column_length_m": col_stats["total_length_m"],
            "column_length_ft_in": col_stats["total_length_ft_in"],
        },
    }


def build_summary(
    beams: list[dict[str, Any]],
    columns: list[dict[str, Any]],
    plates: list[dict[str, Any]],
    bracing: list[dict[str, Any]] | None = None,
    view_metrics: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    bracing = bracing or []
    all_elevs: list[float] = []
    for row in beams:
        for key in ("Start EL", "End EL"):
            v = _parse_numeric_elevation(str(row.get(key, "")))
            if v is not None:
                all_elevs.append(v)
    for row in columns:
        for key in ("Base Elevation", "Top Elevation"):
            v = _parse_numeric_elevation(str(row.get(key, "")))
            if v is not None:
                all_elevs.append(v)
    for row in plates:
        v = _parse_numeric_elevation(str(row.get("Top of Concrete EL", "")))
        if v is not None:
            all_elevs.append(v)

    beam_stats = _total_length_stats(beams, "Length")
    brace_stats = _total_length_stats(bracing, "Length")
    col_stats = _total_length_stats(
        [{**c, "Length": c.get("Height", "")} for c in columns],
        "Length",
    )

    materials: dict[str, int] = {}
    for row in beams + columns + bracing:
        mat = str(row.get("Material", "")).strip()
        if mat:
            materials[mat] = materials.get(mat, 0) + 1

    summary: list[dict[str, Any]] = []

    if view_metrics:
        plan = view_metrics.get("plan") or {}
        elev = view_metrics.get("elevation") or {}
        summary.extend(
            [
                {"Metric": "— PLAN PAGE —", "Value": ""},
                {
                    "Metric": "Plan Pages",
                    "Value": ", ".join(str(p) for p in view_metrics.get("plan_pages") or []) or "N/A",
                },
                {"Metric": "Plan Beam Count", "Value": plan.get("beam_count", 0)},
                {
                    "Metric": "Plan Beam Length",
                    "Value": f"{plan.get('beam_length_ft_in', 'N/A')} ({plan.get('beam_length_m', 0)} m)",
                },
                {"Metric": "Plan Bracing Count", "Value": plan.get("bracing_count", 0)},
                {
                    "Metric": "Plan Bracing Length",
                    "Value": f"{plan.get('bracing_length_ft_in', 'N/A')} ({plan.get('bracing_length_m', 0)} m)",
                },
                {"Metric": "— ELEVATION PAGE —", "Value": ""},
                {
                    "Metric": "Elevation Pages",
                    "Value": ", ".join(str(p) for p in view_metrics.get("elevation_pages") or [])
                    or "N/A",
                },
                {"Metric": "Elevation Column Count", "Value": elev.get("column_count", 0)},
                {
                    "Metric": "Elevation Column Length",
                    "Value": f"{elev.get('column_length_ft_in', 'N/A')} ({elev.get('column_length_m', 0)} m)",
                },
            ]
        )

    summary.extend(
        [
            {"Metric": "— OVERALL —", "Value": ""},
            {"Metric": "Total Beams", "Value": len(beams)},
            {"Metric": "Total Bracing", "Value": len(bracing)},
            {"Metric": "Total Columns", "Value": len(columns)},
            {"Metric": "Total Base Plates", "Value": len(plates)},
            {
                "Metric": "Min Elevation",
                "Value": min(all_elevs) if all_elevs else "N/A",
            },
            {
                "Metric": "Max Elevation",
                "Value": max(all_elevs) if all_elevs else "N/A",
            },
            {
                "Metric": "Total Beam Length (m)",
                "Value": beam_stats["total_length_m"] if beam_stats["total_length_m"] else "N/A",
            },
            {
                "Metric": "Total Bracing Length (m)",
                "Value": brace_stats["total_length_m"] if brace_stats["total_length_m"] else "N/A",
            },
            {
                "Metric": "Total Column Length (m)",
                "Value": col_stats["total_length_m"] if col_stats["total_length_m"] else "N/A",
            },
        ]
    )
    if materials:
        summary.append({"Metric": "— Material Summary —", "Value": ""})
        for mat, count in sorted(materials.items()):
            summary.append({"Metric": f"Material: {mat}", "Value": count})
    else:
        summary.append({"Metric": "Material Summary", "Value": "No materials detected"})
    return summary


def build_excel_bytes(
    beams: list[dict[str, Any]],
    columns: list[dict[str, Any]],
    plates: list[dict[str, Any]],
    bracing: list[dict[str, Any]] | None = None,
    view_metrics: dict[str, Any] | None = None,
) -> bytes:
    bracing = bracing or []
    beam_cols = ["Mark", "Section Size", "Length", "Material", "Start EL", "End EL", "Page"]
    col_cols = [
        "Mark",
        "Section Size",
        "Height",
        "Base Elevation",
        "Top Elevation",
        "Material",
        "Page",
    ]
    plate_cols = [
        "Mark",
        "Plate Size",
        "Thickness",
        "Anchor Bolt Dia",
        "Anchor Bolt Qty",
        "Top of Concrete EL",
        "Page",
    ]
    brace_cols = ["Mark", "Section Size", "Length", "Material", "Page"]

    df_beams = pd.DataFrame(beams, columns=beam_cols) if beams else pd.DataFrame(columns=beam_cols)
    df_cols = pd.DataFrame(columns, columns=col_cols) if columns else pd.DataFrame(columns=col_cols)
    df_plates = (
        pd.DataFrame(plates, columns=plate_cols) if plates else pd.DataFrame(columns=plate_cols)
    )
    df_bracing = (
        pd.DataFrame(bracing, columns=brace_cols) if bracing else pd.DataFrame(columns=brace_cols)
    )
    df_summary = pd.DataFrame(
        build_summary(beams, columns, plates, bracing=bracing, view_metrics=view_metrics)
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_beams.to_excel(writer, sheet_name="Beams", index=False)
        df_cols.to_excel(writer, sheet_name="Columns", index=False)
        df_plates.to_excel(writer, sheet_name="BasePlates", index=False)
        df_bracing.to_excel(writer, sheet_name="Bracing", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
    buffer.seek(0)
    return buffer.read()


# ===========================================================================
# Core pipeline — page by page
# ===========================================================================

def process_pdf(pdf_path: str) -> tuple[bytes, dict[str, Any]]:
    all_beams: list[dict[str, Any]] = []
    all_columns: list[dict[str, Any]] = []
    all_plates: list[dict[str, Any]] = []
    all_bracing: list[dict[str, Any]] = []
    page_sources: list[str] = []
    page_types: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        if total_pages == 0:
            raise HTTPException(status_code=400, detail="PDF has no pages.")

        for index, page in enumerate(pdf.pages):
            page_num = index + 1
            print(f"[Process] Page {page_num}/{total_pages}")

            text, source = extract_page_text(page)
            page_sources.append(source)

            # Supplement with Camelot table text when available
            table_text = extract_tables_text(pdf_path, page_num)
            combined = text
            if table_text:
                combined = f"{text}\n\n--- TABLES ---\n{table_text}"

            page_type = detect_page_type(combined)
            page_types.append({"page": page_num, "type": page_type, "source": source})
            print(f"[Process] Page {page_num} classified as {page_type}")

            # Step 2a: Regex extraction routed by page type
            if page_type == "Plan":
                all_beams.extend(extract_beams_regex(combined, page_num))
                all_beams.extend(extract_plan_beams_from_sections(combined, page_num))
                all_bracing.extend(extract_bracing_regex(combined, page_num))
                all_plates.extend(extract_baseplates_regex(combined, page_num))
            elif page_type == "Elevation":
                all_columns.extend(extract_columns_regex(combined, page_num))
                # Elevations can also show braces; capture if labeled
                all_bracing.extend(extract_bracing_regex(combined, page_num))
            else:
                all_beams.extend(extract_beams_regex(combined, page_num))
                all_columns.extend(extract_columns_regex(combined, page_num))
                all_plates.extend(extract_baseplates_regex(combined, page_num))
                all_bracing.extend(extract_bracing_regex(combined, page_num))

            # Step 2b: AI extraction (if API key present)
            ai_data = ai_extract_from_text(combined, page_num)
            if page_type != "Elevation":
                all_beams.extend(ai_data["beams"])
                all_plates.extend(ai_data["base_plates"])
            if page_type != "Plan":
                all_columns.extend(ai_data["columns"])
            # Bracing from AI if model returns them under beams with BR marks
            for b in ai_data.get("beams") or []:
                mark = str(b.get("Mark", "")).upper()
                if mark.startswith("BR") or mark.startswith("XB"):
                    all_bracing.append(
                        {
                            "Mark": mark,
                            "Section Size": b.get("Section Size", ""),
                            "Length": b.get("Length", ""),
                            "Material": b.get("Material", ""),
                            "Page": page_num,
                        }
                    )

            # Free page resources ASAP for large PDFs
            page.close()

    # Plan section-labeled beams use unique marks; piece-mark beams still merge
    marked_beams = [b for b in all_beams if not str(b.get("Mark", "")).startswith("BM-")]
    section_beams = [b for b in all_beams if str(b.get("Mark", "")).startswith("BM-")]
    beams = _merge_by_mark(marked_beams) + section_beams
    columns = _merge_by_mark(all_columns)
    plates = _merge_by_mark(all_plates)
    bracing = _merge_by_mark(all_bracing)

    view_metrics = build_view_metrics(page_types, beams, bracing, columns)
    excel_bytes = build_excel_bytes(
        beams, columns, plates, bracing=bracing, view_metrics=view_metrics
    )
    meta = {
        "pages": total_pages,
        "beams": beams,
        "columns": columns,
        "base_plates": plates,
        "bracing": bracing,
        "page_types": page_types,
        "view_metrics": view_metrics,
        "summary": build_summary(
            beams, columns, plates, bracing=bracing, view_metrics=view_metrics
        ),
        "page_sources": page_sources,
        "pdf_type": (
            "scanned"
            if page_sources and all(s == "ocr" for s in page_sources)
            else "digital"
            if page_sources and all(s == "digital" for s in page_sources)
            else "mixed"
        ),
    }
    return excel_bytes, meta


# ===========================================================================
# API routes
# ===========================================================================

@app.get("/")
def root():
    return {
        "app": "SteelDraw AI Extractor",
        "status": "ok",
        "upload": "POST /api/upload",
    }


@app.get("/api/health")
def health():
    return {
        "status": "healthy",
        "openai_configured": bool(_openai_api_key()),
    }


@app.post("/api/upload")
async def upload_pdf(file: UploadFile = File(...)):
    """
    Accept a steel drawing PDF (up to 500MB), extract members page-by-page,
    and return an Excel file with Beams / Columns / BasePlates / Summary sheets.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file name provided.")

    filename_lower = file.filename.lower()
    if not filename_lower.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    # Stream to a temp file to avoid loading the entire PDF into RAM
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_path = tmp.name
            total = 0
            chunk_size = 1024 * 1024  # 1 MB chunks
            while True:
                chunk = await file.read(chunk_size)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail="File exceeds the 500MB upload limit.",
                    )
                tmp.write(chunk)

        if total == 0:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        excel_bytes, meta = process_pdf(tmp_path)

        # Also expose JSON preview via custom headers (small counts) —
        # the Excel is the primary download payload.
        download_name = Path(file.filename).stem + "_SteelDraw_Extract.xlsx"
        headers = {
            "Content-Disposition": f'attachment; filename="{download_name}"',
            "X-SteelDraw-Pages": str(meta["pages"]),
            "X-SteelDraw-Beams": str(len(meta["beams"])),
            "X-SteelDraw-Columns": str(len(meta["columns"])),
            "X-SteelDraw-BasePlates": str(len(meta["base_plates"])),
            "X-SteelDraw-PdfType": str(meta["pdf_type"]),
            "Access-Control-Expose-Headers": (
                "Content-Disposition, X-SteelDraw-Pages, X-SteelDraw-Beams, "
                "X-SteelDraw-Columns, X-SteelDraw-BasePlates, X-SteelDraw-PdfType"
            ),
        }
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"Failed to process PDF: {exc}",
        ) from exc
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


@app.post("/api/extract")
async def extract_preview(file: UploadFile = File(...)):
    """
    Same pipeline as /api/upload but returns JSON preview for the UI table,
    plus a base64 Excel payload for download.
    """
    import base64

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_path = tmp.name
            total = 0
            chunk_size = 1024 * 1024
            while True:
                chunk = await file.read(chunk_size)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail="File exceeds the 500MB upload limit.",
                    )
                tmp.write(chunk)

        if total == 0:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        excel_bytes, meta = process_pdf(tmp_path)
        download_name = Path(file.filename).stem + "_SteelDraw_Extract.xlsx"

        return {
            "filename": download_name,
            "pages": meta["pages"],
            "pdf_type": meta["pdf_type"],
            "page_types": meta.get("page_types", []),
            "view_metrics": meta.get("view_metrics", {}),
            "beams": meta["beams"],
            "columns": meta["columns"],
            "base_plates": meta["base_plates"],
            "bracing": meta.get("bracing", []),
            "summary": meta["summary"],
            "excel_base64": base64.b64encode(excel_bytes).decode("ascii"),
        }
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"Failed to process PDF: {exc}",
        ) from exc
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
