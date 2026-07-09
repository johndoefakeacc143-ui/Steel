"""
SteelDraw AI Extractor — FastAPI Backend
=========================================
Upload steel structure engineering PDFs, extract Beams / Columns / Base Plates
(and Bracings for Summary), then return a multi-sheet Excel workbook.

Edit the REGEX PATTERNS section below when you need to tune mark/size matching.
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import uuid
from collections import Counter
from pathlib import Path
from typing import Any, Optional

import cv2
import numpy as np
import pandas as pd
import pdfplumber
import pytesseract
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Env / App setup
# ---------------------------------------------------------------------------
load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")
load_dotenv()  # also allow backend/.env

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
MAX_UPLOAD_BYTES = 500 * 1024 * 1024  # 500 MB
PAGE_ASK_THRESHOLD = 5  # >5 pages → ask user which pages to scan

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

# In-memory job store for multi-step uploads (page selection → extract)
# Keys: job_id → { pdf_path, page_count, filename }
JOBS: dict[str, dict[str, Any]] = {}


# =============================================================================
# REGEX PATTERNS — edit these to match your drawing conventions
# =============================================================================
#
# Beam marks: B1, B-12, BM3, Beam 4, etc.
BEAM_MARK_RE = re.compile(
    r"\b(?:B|BM|BEAM)[-\s]?(\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

# Column marks: C1, C-12, COL3, Column 4, etc.
COLUMN_MARK_RE = re.compile(
    r"\b(?:C|COL|COLUMN)[-\s]?(\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

# Base plate marks: BP1, BP-12, BASE PLATE 3, etc.
BASE_PLATE_MARK_RE = re.compile(
    r"\b(?:BP|B\.?P\.?|BASE\s*PLATE)[-\s]?(\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

# Bracing marks: BR1, BR-4, BRACING 2, etc.
BRACING_MARK_RE = re.compile(
    r"\b(?:BR|BRG|BRACING)[-\s]?(\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

# Section sizes — W-shapes, ISMB/ISMC/ISMB, UB/UC, HSS, pipes, angles, plates
# Examples: W18x35, W21×44, ISMB400, ISMC300, UB457x191x67, HSS6x6x1/4, L4x4x3/8
SECTION_SIZE_RE = re.compile(
    r"\b(?:"
    r"W\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # W18x35
    r"|ISMB\s*\d{2,4}"  # ISMB400
    r"|ISMC\s*\d{2,4}"  # ISMC300
    r"|ISMB\s*\d{2,4}"  # duplicate-safe
    r"|UB\s*\d{2,4}\s*[x×]\s*\d{2,4}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # UB457x191x67
    r"|UC\s*\d{2,4}\s*[x×]\s*\d{2,4}\s*[x×]\s*\d{1,3}(?:\.\d+)?"  # UC
    r"|HSS\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d/]+"  # HSS
    r"|PIPE\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d.]+"  # PIPE
    r"|L\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d./]+"  # Angle
    r"|PL\s*\d+(?:\.\d+)?\s*[x×]\s*\d+(?:\.\d+)?"  # Plate section
    r"|RHS\s*\d+\s*[x×]\s*\d+\s*[x×]\s*\d+"  # RHS
    r"|SHS\s*\d+\s*[x×]\s*\d+"  # SHS
    r")\b",
    re.IGNORECASE,
)

# Base plate size: 600x600x30, 500×500×25, etc. (LxWxThk in mm)
BASE_PLATE_SIZE_RE = re.compile(
    r"\b(\d{2,4})\s*[x×]\s*(\d{2,4})\s*[x×]\s*(\d{1,3}(?:\.\d+)?)\b",
    re.IGNORECASE,
)

# Length / height dimensions (mm or m). Prefer values near marks.
# Examples: 6000, 6.0m, 4500 mm, L=3200
LENGTH_RE = re.compile(
    r"(?:L\s*[=:]?\s*)?(\d{3,5}(?:\.\d+)?)\s*(?:mm)?\b"
    r"|(?:L\s*[=:]?\s*)?(\d{1,2}(?:\.\d+)?)\s*m\b",
    re.IGNORECASE,
)

# Elevations: EL +5000, EL. 12.500, Elevation +0, TOC +0.00
# Avoid matching inside "LEVEL" by requiring EL as a whole token or ELEV/ELEVATION.
ELEVATION_RE = re.compile(
    r"(?<![A-Z])(?:EL\.?|ELEV(?:ATION)?\.?|TOC|TOG|TOS)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)",
    re.IGNORECASE,
)

# Explicit base / top elevation phrases on column schedules
BASE_ELEV_RE = re.compile(
    r"(?:BASE(?:\s+OF)?\s*(?:EL(?:EV(?:ATION)?)?\.?)?|B\.?\s*EL\.?)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)",
    re.IGNORECASE,
)
TOP_ELEV_RE = re.compile(
    r"(?:TOP(?:\s+OF)?\s*(?:EL(?:EV(?:ATION)?)?\.?)?|T\.?\s*EL\.?)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)",
    re.IGNORECASE,
)

# Material grades: A36, A992, S275, S355, Fe410, GR.50, ASTM A36
MATERIAL_RE = re.compile(
    r"\b(?:"
    r"A\d{2,3}(?:M)?"  # A36, A992
    r"|S\d{3}"  # S275, S355
    r"|Fe\d{3}"  # Fe410
    r"|GR\.?\s*\d{2}"  # GR.50
    r"|ASTM\s*A\d{2,3}"
    r"|EN\s*10025"
    r")\b",
    re.IGNORECASE,
)

# Anchor bolts: 4-M20, 4Ø20, 4-#3/4", M24 x 4, 6 Nos M20
ANCHOR_BOLT_RE = re.compile(
    r"\b(\d+)\s*(?:[-x×]|Nos?\.?\s*)?(?:M|#|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?|\d/\d)\s*(?:\"|mm)?\b"
    r"|\b(M\d{1,2})\s*[x×]\s*(\d+)\b",
    re.IGNORECASE,
)

# Weight for base plates (kg): 45 kg, WT=52.3, Weight 120kg
WEIGHT_RE = re.compile(
    r"(?:WT|WEIGHT|MASS)\s*[=:]?\s*(\d+(?:\.\d+)?)\s*(?:kg|kgs|KG)?"
    r"|(\d+(?:\.\d+)?)\s*(?:kg|kgs)\b",
    re.IGNORECASE,
)


# =============================================================================
# PDF helpers — digital vs scanned, page text, OCR
# =============================================================================

def get_page_count(pdf_path: str) -> int:
    """Return number of pages without loading the whole file into memory."""
    reader = PdfReader(pdf_path)
    return len(reader.pages)


def is_page_digital(page: pdfplumber.page.Page, min_chars: int = 40) -> bool:
    """
    Heuristic: if a page has enough extractable characters, treat as digital.
    Scanned pages typically return little/no text via pdfplumber.
    """
    text = page.extract_text() or ""
    # Strip whitespace-only noise
    cleaned = re.sub(r"\s+", "", text)
    return len(cleaned) >= min_chars


def preprocess_for_ocr(image_bgr: np.ndarray) -> np.ndarray:
    """
    OpenCV preprocessing for scanned drawing OCR:
    grayscale → denoise → adaptive threshold → slight dilation for thin lines.
    """
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    # Bilateral filter keeps edges while reducing noise on large drawings
    denoised = cv2.bilateralFilter(gray, d=5, sigmaColor=50, sigmaSpace=50)
    # Adaptive threshold handles uneven scan lighting
    binary = cv2.adaptiveThreshold(
        denoised,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        11,
    )
    # Invert if mostly black (OCR prefers dark text on light bg)
    if np.mean(binary) < 127:
        binary = cv2.bitwise_not(binary)
    kernel = np.ones((1, 1), np.uint8)
    return cv2.dilate(binary, kernel, iterations=1)


def ocr_page_image(page: pdfplumber.page.Page, dpi: int = 200) -> str:
    """Rasterize a pdfplumber page and run Tesseract OCR."""
    # Lower DPI for huge sheets to avoid OOM; raise if marks are missed
    pil_img = page.to_image(resolution=dpi).original
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    processed = preprocess_for_ocr(bgr)
    # --psm 6: assume a single uniform block of text (good for drawings + notes)
    config = "--psm 6 -c preserve_interword_spaces=1"
    return pytesseract.image_to_string(processed, config=config) or ""


def extract_tables_camelot(pdf_path: str, page_no: int) -> str:
    """
    Optional Camelot table pass — useful when member schedules are tabular.
    Returns concatenated table text, or '' if Camelot/Ghostscript unavailable.
    """
    try:
        import camelot

        # flavor='lattice' for ruled schedule tables; fall back to stream
        tables = camelot.read_pdf(
            pdf_path, pages=str(page_no), flavor="lattice", suppress_stdout=True
        )
        if tables.n == 0:
            tables = camelot.read_pdf(
                pdf_path, pages=str(page_no), flavor="stream", suppress_stdout=True
            )
        chunks = []
        for t in tables:
            try:
                chunks.append(t.df.to_string(index=False, header=False))
            except Exception:  # noqa: BLE001
                continue
        return "\n".join(chunks)
    except Exception as exc:  # noqa: BLE001
        print(f"[Camelot skip page {page_no}] {exc}")
        return ""


def extract_page_text(
    page: pdfplumber.page.Page,
    pdf_path: str,
    page_no: int,
) -> tuple[str, str]:
    """
    Return (text, source) where source is 'digital' or 'ocr'.
    Process one page at a time to keep memory bounded.
    Appends Camelot table text when available (member schedules).
    """
    table_text = extract_tables_camelot(pdf_path, page_no)
    if is_page_digital(page):
        text = page.extract_text() or ""
        if table_text:
            text = f"{text}\n\n{table_text}"
        return (text, "digital")
    try:
        text = ocr_page_image(page)
        if table_text:
            text = f"{text}\n\n{table_text}"
        return (text, "ocr")
    except Exception as exc:  # noqa: BLE001 — fall back gracefully
        print(f"[OCR warning] {exc}")
        text = page.extract_text() or ""
        if table_text:
            text = f"{text}\n\n{table_text}"
        return (text, "digital-fallback")


# =============================================================================
# Regex extraction (primary) — works without OpenAI
# =============================================================================

def _normalize_section(raw: str) -> str:
    return re.sub(r"\s+", "", raw.upper().replace("×", "x"))


def _parse_length_mm(match: re.Match) -> Optional[float]:
    """Convert LENGTH_RE match groups to millimetres."""
    if match.group(1):
        return float(match.group(1))
    if match.group(2):
        return float(match.group(2)) * 1000.0
    return None


def _nearby_window(text: str, start: int, end: int, radius: int = 180) -> str:
    """
    Prefer the full line containing the mark (schedule-style drawings),
    then fall back to a character window for free-floating callouts.
    """
    line_start = text.rfind("\n", 0, start) + 1
    line_end = text.find("\n", end)
    if line_end == -1:
        line_end = len(text)
    line = text[line_start:line_end].strip()
    # If the line is informative enough, use it (+ maybe next line for wrapped notes)
    if len(line) >= 8:
        next_end = text.find("\n", line_end + 1)
        if next_end == -1:
            next_end = min(len(text), line_end + 120)
        nxt = text[line_end:next_end].strip()
        # Only append next line if it looks like a continuation (no new mark)
        if nxt and not BEAM_MARK_RE.search(nxt) and not COLUMN_MARK_RE.search(nxt):
            if not BASE_PLATE_MARK_RE.search(nxt) and not BRACING_MARK_RE.search(nxt):
                return f"{line} {nxt}"
        return line
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    return text[lo:hi]


def _first_section(window: str) -> str:
    m = SECTION_SIZE_RE.search(window)
    return _normalize_section(m.group(0)) if m else ""


def _first_length(window: str) -> Optional[float]:
    # Prefer explicit L= / Length / Height patterns first
    explicit = re.search(
        r"(?:L|LEN(?:GTH)?|H(?:EIGHT)?|HT)\s*[=:]?\s*(\d{3,5}(?:\.\d+)?)\s*(?:mm)?"
        r"|(?:L|LEN(?:GTH)?|H(?:EIGHT)?|HT)\s*[=:]?\s*(\d{1,2}(?:\.\d+)?)\s*m\b",
        window,
        re.IGNORECASE,
    )
    if explicit:
        if explicit.group(1):
            return float(explicit.group(1))
        if explicit.group(2):
            return float(explicit.group(2)) * 1000.0
    for m in LENGTH_RE.finditer(window):
        val = _parse_length_mm(m)
        # Skip tiny numbers that are likely grid refs / bolt sizes
        if val is not None and val >= 200:
            return val
    return None


def _elevations(window: str) -> list[float]:
    vals = []
    for m in ELEVATION_RE.finditer(window):
        try:
            vals.append(float(m.group(1)))
        except ValueError:
            continue
    return vals


def _first_material(window: str) -> str:
    m = MATERIAL_RE.search(window)
    return m.group(0).upper().replace(" ", "") if m else ""


def extract_beams_regex(text: str, page_no: int) -> list[dict[str, Any]]:
    """Extract beam rows from page text using mark-centric regex."""
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in BEAM_MARK_RE.finditer(text):
        mark = f"B{m.group(1).upper()}"
        if mark in seen:
            continue
        seen.add(mark)
        window = _nearby_window(text, m.start(), m.end())
        elevs = _elevations(window)
        length = _first_length(window)
        start_el: Any = ""
        end_el: Any = ""
        start_m = re.search(
            r"START\s*(?:EL(?:EV(?:ATION)?)?\.?)?\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)",
            window,
            re.IGNORECASE,
        )
        end_m = re.search(
            r"END\s*(?:EL(?:EV(?:ATION)?)?\.?)?\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)",
            window,
            re.IGNORECASE,
        )
        if start_m:
            start_el = float(start_m.group(1))
        if end_m:
            end_el = float(end_m.group(1))
        if start_el == "" and elevs:
            start_el = elevs[0]
        if end_el == "":
            end_el = elevs[1] if len(elevs) >= 2 else (elevs[0] if elevs else "")
        rows.append(
            {
                "Mark": mark,
                "Section Size": _first_section(window),
                "Length (mm)": length if length is not None else "",
                "Material": _first_material(window) or "A36",
                "Start EL": start_el,
                "End EL": end_el,
                "Page": page_no,
            }
        )
    return rows


def extract_columns_regex(text: str, page_no: int) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in COLUMN_MARK_RE.finditer(text):
        mark = f"C{m.group(1).upper()}"
        if mark in seen:
            continue
        seen.add(mark)
        window = _nearby_window(text, m.start(), m.end())
        elevs = _elevations(window)
        base_m = BASE_ELEV_RE.search(window)
        top_m = TOP_ELEV_RE.search(window)
        base_el: Any = float(base_m.group(1)) if base_m else (elevs[0] if elevs else "")
        top_el: Any = float(top_m.group(1)) if top_m else (elevs[-1] if len(elevs) >= 2 else "")
        height = _first_length(window)
        # Prefer height from top-base elevation delta when available
        try:
            if base_el != "" and top_el != "" and height is None:
                height = abs(float(top_el) - float(base_el))
        except (TypeError, ValueError):
            pass
        rows.append(
            {
                "Mark": mark,
                "Section Size": _first_section(window),
                "Height (mm)": height if height is not None else "",
                "Base Elevation": base_el,
                "Top Elevation": top_el,
                "Material": _first_material(window) or "A992",
                "Page": page_no,
            }
        )
    return rows


def extract_base_plates_regex(text: str, page_no: int) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in BASE_PLATE_MARK_RE.finditer(text):
        mark = f"BP{m.group(1).upper()}"
        if mark in seen:
            continue
        seen.add(mark)
        window = _nearby_window(text, m.start(), m.end(), radius=220)
        size_m = BASE_PLATE_SIZE_RE.search(window)
        plate_size = ""
        thickness = ""
        if size_m:
            plate_size = f"{size_m.group(1)}x{size_m.group(2)}x{size_m.group(3)}"
            thickness = size_m.group(3)

        bolt_dia, bolt_qty = "", ""
        ab = ANCHOR_BOLT_RE.search(window)
        if ab:
            if ab.group(1) and ab.group(2):
                bolt_qty, bolt_dia = ab.group(1), f"M{ab.group(2)}" if not str(ab.group(2)).startswith("M") else ab.group(2)
                if not str(bolt_dia).upper().startswith("M") and "/" not in str(bolt_dia):
                    bolt_dia = f"M{bolt_dia}"
            elif ab.group(3) and ab.group(4):
                bolt_dia, bolt_qty = ab.group(3).upper(), ab.group(4)

        elevs = _elevations(window)
        wt = ""
        wm = WEIGHT_RE.search(window)
        if wm:
            wt = wm.group(1) or wm.group(2) or ""
        # Estimate weight (kg) from plate size if missing: L*W*T * 7.85e-6
        if not wt and size_m:
            try:
                l, w, t = float(size_m.group(1)), float(size_m.group(2)), float(size_m.group(3))
                wt = round(l * w * t * 7.85e-6, 2)
            except ValueError:
                pass

        rows.append(
            {
                "Mark": mark,
                "Plate Size": plate_size,
                "Thickness (mm)": thickness,
                "Anchor Bolt Dia": bolt_dia,
                "Anchor Bolt Qty": bolt_qty,
                "Top of Concrete EL": elevs[0] if elevs else "",
                "Weight (kg)": wt,
                "Page": page_no,
            }
        )
    return rows


def extract_bracings_regex(text: str, page_no: int) -> list[dict[str, Any]]:
    """Bracings feed the Summary sheet (count by length)."""
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in BRACING_MARK_RE.finditer(text):
        mark = f"BR{m.group(1).upper()}"
        if mark in seen:
            continue
        seen.add(mark)
        window = _nearby_window(text, m.start(), m.end())
        length = _first_length(window)
        rows.append(
            {
                "Mark": mark,
                "Section Size": _first_section(window),
                "Length (mm)": length if length is not None else "",
                "Page": page_no,
            }
        )
    return rows


# =============================================================================
# Optional AI enrichment via LangChain + OpenAI
# =============================================================================

AI_SYSTEM_PROMPT = """You are a senior steel detailer / structural steel modeler.
Given OCR or digital text from a steel structure engineering drawing page,
extract structural members as JSON with keys: beams, columns, base_plates, bracings.

beams: [{mark, section_size, length_mm, material, start_el, end_el}]
columns: [{mark, section_size, height_mm, base_elevation, top_elevation, material}]
base_plates: [{mark, plate_size, thickness_mm, anchor_bolt_dia, anchor_bolt_qty, top_of_concrete_el, weight_kg}]
bracings: [{mark, section_size, length_mm}]

Rules:
- Use marks exactly as on the drawing (B1, C3, BP2, BR1).
- Section sizes like W18x35, ISMB400, UB457x191x67.
- Lengths/heights in millimetres (convert metres ×1000).
- If a field is unknown, use empty string.
- Return ONLY valid JSON, no markdown.
"""


def ai_extract_from_text(text: str, page_no: int) -> dict[str, list[dict]]:
    """
    Call OpenAI (via langchain) to enrich extraction.
    Returns empty lists if no API key or on failure — regex results still apply.
    """
    empty = {"beams": [], "columns": [], "base_plates": [], "bracings": []}
    if not OPENAI_API_KEY or not text.strip():
        return empty
    try:
        from langchain_openai import ChatOpenAI
        from langchain_core.messages import HumanMessage, SystemMessage

        llm = ChatOpenAI(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            temperature=0,
            api_key=OPENAI_API_KEY,
        )
        # Cap text to keep token usage reasonable on huge sheets
        clipped = text[:12000]
        resp = llm.invoke(
            [
                SystemMessage(content=AI_SYSTEM_PROMPT),
                HumanMessage(
                    content=f"Page {page_no} drawing text:\n\n{clipped}\n\nExtract members as JSON."
                ),
            ]
        )
        content = resp.content if isinstance(resp.content, str) else str(resp.content)
        # Strip optional ```json fences
        content = re.sub(r"^```(?:json)?\s*", "", content.strip())
        content = re.sub(r"\s*```$", "", content)
        import json

        data = json.loads(content)
        return {
            "beams": data.get("beams") or [],
            "columns": data.get("columns") or [],
            "base_plates": data.get("base_plates") or data.get("basePlates") or [],
            "bracings": data.get("bracings") or [],
        }
    except Exception as exc:  # noqa: BLE001
        print(f"[AI extract warning page {page_no}] {exc}")
        return empty


def _merge_ai_beams(ai_rows: list[dict], page_no: int) -> list[dict[str, Any]]:
    out = []
    for r in ai_rows:
        mark = str(r.get("mark") or "").upper().replace(" ", "")
        if not mark:
            continue
        if not mark.startswith("B"):
            mark = f"B{mark}"
        out.append(
            {
                "Mark": mark,
                "Section Size": _normalize_section(str(r.get("section_size") or "")),
                "Length (mm)": r.get("length_mm") or "",
                "Material": r.get("material") or "",
                "Start EL": r.get("start_el") or "",
                "End EL": r.get("end_el") or "",
                "Page": page_no,
            }
        )
    return out


def _merge_ai_columns(ai_rows: list[dict], page_no: int) -> list[dict[str, Any]]:
    out = []
    for r in ai_rows:
        mark = str(r.get("mark") or "").upper().replace(" ", "")
        if not mark:
            continue
        if not mark.startswith("C"):
            mark = f"C{mark}"
        out.append(
            {
                "Mark": mark,
                "Section Size": _normalize_section(str(r.get("section_size") or "")),
                "Height (mm)": r.get("height_mm") or "",
                "Base Elevation": r.get("base_elevation") or "",
                "Top Elevation": r.get("top_elevation") or "",
                "Material": r.get("material") or "",
                "Page": page_no,
            }
        )
    return out


def _merge_ai_base_plates(ai_rows: list[dict], page_no: int) -> list[dict[str, Any]]:
    out = []
    for r in ai_rows:
        mark = str(r.get("mark") or "").upper().replace(" ", "")
        if not mark:
            continue
        if not mark.startswith("BP"):
            mark = f"BP{mark}" if not mark.startswith("B") else mark
        out.append(
            {
                "Mark": mark,
                "Plate Size": str(r.get("plate_size") or "").replace("×", "x"),
                "Thickness (mm)": r.get("thickness_mm") or "",
                "Anchor Bolt Dia": r.get("anchor_bolt_dia") or "",
                "Anchor Bolt Qty": r.get("anchor_bolt_qty") or "",
                "Top of Concrete EL": r.get("top_of_concrete_el") or "",
                "Weight (kg)": r.get("weight_kg") or "",
                "Page": page_no,
            }
        )
    return out


def _merge_ai_bracings(ai_rows: list[dict], page_no: int) -> list[dict[str, Any]]:
    out = []
    for r in ai_rows:
        mark = str(r.get("mark") or "").upper().replace(" ", "")
        if not mark:
            continue
        if not mark.startswith("BR"):
            mark = f"BR{mark}"
        out.append(
            {
                "Mark": mark,
                "Section Size": _normalize_section(str(r.get("section_size") or "")),
                "Length (mm)": r.get("length_mm") or "",
                "Page": page_no,
            }
        )
    return out


def dedupe_by_mark(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Keep the richest row per Mark (most non-empty fields)."""
    best: dict[str, dict[str, Any]] = {}
    for row in rows:
        mark = row.get("Mark") or ""
        if not mark:
            continue
        score = sum(1 for v in row.values() if v not in ("", None))
        if mark not in best or score > sum(1 for v in best[mark].values() if v not in ("", None)):
            best[mark] = row
    return list(best.values())


# =============================================================================
# Page-by-page pipeline
# =============================================================================

def parse_page_list(raw: Optional[str], total_pages: int) -> list[int]:
    """
    Parse '1,3,5-7' style page lists (1-based) into sorted unique page numbers.
    Empty / None → all pages.
    """
    if not raw or not str(raw).strip():
        return list(range(1, total_pages + 1))
    pages: set[int] = set()
    for part in str(raw).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start, end = int(a.strip()), int(b.strip())
            for p in range(min(start, end), max(start, end) + 1):
                if 1 <= p <= total_pages:
                    pages.add(p)
        else:
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p)
    return sorted(pages) if pages else list(range(1, total_pages + 1))


def process_pdf(
    pdf_path: str,
    plan_pages: Optional[list[int]] = None,
    elevation_pages: Optional[list[int]] = None,
) -> dict[str, Any]:
    """
    Process PDF page-by-page.
    - plan_pages: used for beams + bracings
    - elevation_pages: used for columns + base plates
    If either is None, scan those member types on all selected/all pages.
    """
    beams: list[dict] = []
    columns: list[dict] = []
    base_plates: list[dict] = []
    bracings: list[dict] = []
    page_sources: dict[int, str] = {}
    all_text_snippets: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        plan_set = set(plan_pages) if plan_pages is not None else set(range(1, total + 1))
        elev_set = set(elevation_pages) if elevation_pages is not None else set(range(1, total + 1))
        pages_needed = sorted(plan_set | elev_set)

        for page_no in pages_needed:
            page = pdf.pages[page_no - 1]
            text, source = extract_page_text(page, pdf_path, page_no)
            page_sources[page_no] = source
            all_text_snippets.append(f"--- Page {page_no} ---\n{text[:2000]}")

            # Regex pass
            if page_no in plan_set:
                beams.extend(extract_beams_regex(text, page_no))
                bracings.extend(extract_bracings_regex(text, page_no))
            if page_no in elev_set:
                columns.extend(extract_columns_regex(text, page_no))
                base_plates.extend(extract_base_plates_regex(text, page_no))

            # AI enrichment pass (optional)
            ai = ai_extract_from_text(text, page_no)
            if page_no in plan_set:
                beams.extend(_merge_ai_beams(ai["beams"], page_no))
                bracings.extend(_merge_ai_bracings(ai["bracings"], page_no))
            if page_no in elev_set:
                columns.extend(_merge_ai_columns(ai["columns"], page_no))
                base_plates.extend(_merge_ai_base_plates(ai["base_plates"], page_no))

            # Free page resources promptly
            del page

    beams = dedupe_by_mark(beams)
    columns = dedupe_by_mark(columns)
    base_plates = dedupe_by_mark(base_plates)
    bracings = dedupe_by_mark(bracings)

    summary_rows, engineer_notes = build_summary(
        beams, columns, base_plates, bracings, page_sources, all_text_snippets
    )

    return {
        "beams": beams,
        "columns": columns,
        "base_plates": base_plates,
        "bracings": bracings,
        "summary": summary_rows,
        "engineer_notes": engineer_notes,
        "page_sources": page_sources,
    }


# =============================================================================
# Summary — senior steel engineer narrative + counts
# =============================================================================

def _count_by_key(rows: list[dict], key: str) -> list[tuple[str, int]]:
    counter: Counter = Counter()
    for r in rows:
        val = r.get(key)
        if val in ("", None):
            val = "UNKNOWN"
        counter[str(val)] += 1
    return sorted(counter.items(), key=lambda x: (-x[1], x[0]))


def _numeric_series(rows: list[dict], key: str) -> list[float]:
    vals = []
    for r in rows:
        v = r.get(key)
        if v in ("", None):
            continue
        try:
            vals.append(float(v))
        except (TypeError, ValueError):
            continue
    return vals


def build_summary(
    beams: list[dict],
    columns: list[dict],
    base_plates: list[dict],
    bracings: list[dict],
    page_sources: dict[int, str],
    text_snippets: list[str],
) -> tuple[list[dict[str, Any]], str]:
    """
    Sheet4 Summary content:
    - Totals, min/max elevation, total beam length, material summary
    - Engineer-style counts: N columns of length L, N bracings of length L, etc.
    """
    rows: list[dict[str, Any]] = []

    def add(category: str, metric: str, value: Any) -> None:
        rows.append({"Category": category, "Metric": metric, "Value": value})

    add("Totals", "Beam Count", len(beams))
    add("Totals", "Column Count", len(columns))
    add("Totals", "Base Plate Count", len(base_plates))
    add("Totals", "Bracing Count", len(bracings))

    # Beam length totals
    beam_lens = _numeric_series(beams, "Length (mm)")
    add("Beams", "Total Beam Length (mm)", round(sum(beam_lens), 1) if beam_lens else 0)
    add("Beams", "Total Beam Length (m)", round(sum(beam_lens) / 1000, 2) if beam_lens else 0)

    # Elevations across members
    elevs: list[float] = []
    for r in beams:
        elevs.extend(_numeric_series([r], "Start EL"))
        elevs.extend(_numeric_series([r], "End EL"))
    for r in columns:
        elevs.extend(_numeric_series([r], "Base Elevation"))
        elevs.extend(_numeric_series([r], "Top Elevation"))
    for r in base_plates:
        elevs.extend(_numeric_series([r], "Top of Concrete EL"))
    add("Elevations", "Min Elevation", min(elevs) if elevs else "")
    add("Elevations", "Max Elevation", max(elevs) if elevs else "")

    # Size counts
    for size, cnt in _count_by_key(beams, "Section Size"):
        add("Beam Sizes", f"{cnt} beam(s) of size {size}", cnt)
    for size, cnt in _count_by_key(columns, "Section Size"):
        add("Column Sizes", f"{cnt} column(s) of size {size}", cnt)
    for size, cnt in _count_by_key(base_plates, "Plate Size"):
        add("Base Plate Sizes", f"{cnt} base plate(s) of size {size}", cnt)

    # Length / weight groupings (engineer narrative inputs)
    for length, cnt in _count_by_key(columns, "Height (mm)"):
        add("Column Lengths", f"{cnt} column(s) of length {length} mm", cnt)
    for length, cnt in _count_by_key(bracings, "Length (mm)"):
        add("Bracing Lengths", f"{cnt} bracing(s) of length {length} mm", cnt)
    for length, cnt in _count_by_key(beams, "Length (mm)"):
        add("Beam Lengths", f"{cnt} beam(s) of length {length} mm", cnt)
    for wt, cnt in _count_by_key(base_plates, "Weight (kg)"):
        add("Base Plate Weights", f"{cnt} base plate(s) of weight {wt} kg", cnt)

    # Material summary
    mat_counter: Counter = Counter()
    for r in beams + columns:
        mat = r.get("Material") or "UNKNOWN"
        mat_counter[str(mat)] += 1
    for mat, cnt in sorted(mat_counter.items(), key=lambda x: (-x[1], x[0])):
        add("Material Summary", mat, cnt)

    # Page source mix
    src_counter = Counter(page_sources.values())
    for src, cnt in src_counter.items():
        add("PDF Analysis", f"Pages processed as {src}", cnt)

    # Engineer narrative
    notes_lines = [
        "SENIOR STEEL STRUCTURE ENGINEER — DRAWING SUMMARY",
        "=" * 56,
        f"This drawing set contains {len(columns)} column(s), {len(beams)} beam(s), "
        f"{len(bracings)} bracing(s), and {len(base_plates)} base plate(s).",
        "",
    ]
    if columns:
        notes_lines.append("Columns by height:")
        for length, cnt in _count_by_key(columns, "Height (mm)"):
            notes_lines.append(f"  • {cnt} column(s) of length {length} mm")
        notes_lines.append("Columns by section:")
        for size, cnt in _count_by_key(columns, "Section Size"):
            notes_lines.append(f"  • {cnt} column(s) of size {size}")
        notes_lines.append("")
    if bracings:
        notes_lines.append("Bracings by length:")
        for length, cnt in _count_by_key(bracings, "Length (mm)"):
            notes_lines.append(f"  • {cnt} bracing(s) of length {length} mm")
        notes_lines.append("")
    if beams:
        notes_lines.append("Beams by length:")
        for length, cnt in _count_by_key(beams, "Length (mm)"):
            notes_lines.append(f"  • {cnt} beam(s) of length {length} mm")
        notes_lines.append("Beams by section:")
        for size, cnt in _count_by_key(beams, "Section Size"):
            notes_lines.append(f"  • {cnt} beam(s) of size {size}")
        notes_lines.append("")
    if base_plates:
        notes_lines.append("Base plates by weight:")
        for wt, cnt in _count_by_key(base_plates, "Weight (kg)"):
            notes_lines.append(f"  • {cnt} base plate(s) of weight {wt} kg")
        notes_lines.append("Base plates by size:")
        for size, cnt in _count_by_key(base_plates, "Plate Size"):
            notes_lines.append(f"  • {cnt} base plate(s) of size {size}")
        notes_lines.append("")

    if elevs:
        notes_lines.append(
            f"Elevation range observed: {min(elevs)} to {max(elevs)} "
            "(verify against grid/EL callouts on the sheets)."
        )
    if beam_lens:
        notes_lines.append(
            f"Aggregate beam length: {round(sum(beam_lens)/1000, 2)} m "
            f"({round(sum(beam_lens), 1)} mm)."
        )
    if mat_counter:
        mats = ", ".join(f"{m} ({c})" for m, c in mat_counter.most_common())
        notes_lines.append(f"Material mix: {mats}.")

    notes_lines.extend(
        [
            "",
            "Recommendation: Cross-check marks against the member schedule / BOM "
            "and confirm start/end elevations on the elevation sheets before fabrication.",
        ]
    )

    # Optional AI polish of the narrative
    engineer_notes = "\n".join(notes_lines)
    if OPENAI_API_KEY:
        try:
            from langchain_openai import ChatOpenAI
            from langchain_core.messages import HumanMessage, SystemMessage

            llm = ChatOpenAI(
                model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
                temperature=0.2,
                api_key=OPENAI_API_KEY,
            )
            polish = llm.invoke(
                [
                    SystemMessage(
                        content=(
                            "You are a senior structural steel engineer. "
                            "Rewrite the notes into a concise professional summary "
                            "(keep all counts and lengths). Plain text only."
                        )
                    ),
                    HumanMessage(content=engineer_notes[:8000]),
                ]
            )
            polished = polish.content if isinstance(polish.content, str) else str(polish.content)
            if polished.strip():
                engineer_notes = polished.strip()
        except Exception as exc:  # noqa: BLE001
            print(f"[AI summary warning] {exc}")

    add("Engineer Notes", "Narrative", engineer_notes)
    return rows, engineer_notes


# =============================================================================
# Excel workbook
# =============================================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")


def _write_sheet(ws, rows: list[dict], columns: list[str]) -> None:
    df = pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(columns=columns)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        max_len = 0
        letter = col[0].column_letter
        for cell in col:
            max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 60)


def build_excel(result: dict[str, Any]) -> bytes:
    wb = Workbook()

    # Sheet 1 — Beams
    ws1 = wb.active
    ws1.title = "Beams"
    _write_sheet(
        ws1,
        result["beams"],
        ["Mark", "Section Size", "Length (mm)", "Material", "Start EL", "End EL", "Page"],
    )

    # Sheet 2 — Columns
    ws2 = wb.create_sheet("Columns")
    _write_sheet(
        ws2,
        result["columns"],
        [
            "Mark",
            "Section Size",
            "Height (mm)",
            "Base Elevation",
            "Top Elevation",
            "Material",
            "Page",
        ],
    )

    # Sheet 3 — BasePlates
    ws3 = wb.create_sheet("BasePlates")
    _write_sheet(
        ws3,
        result["base_plates"],
        [
            "Mark",
            "Plate Size",
            "Thickness (mm)",
            "Anchor Bolt Dia",
            "Anchor Bolt Qty",
            "Top of Concrete EL",
            "Weight (kg)",
            "Page",
        ],
    )

    # Sheet 4 — Summary
    ws4 = wb.create_sheet("Summary")
    _write_sheet(ws4, result["summary"], ["Category", "Metric", "Value"])
    # Append full narrative at the bottom for readability
    start = len(result["summary"]) + 3
    ws4.cell(row=start, column=1, value="Engineer Narrative").font = Font(bold=True)
    cell = ws4.cell(row=start + 1, column=1, value=result.get("engineer_notes", ""))
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.fill = NOTE_FILL
    ws4.merge_cells(start_row=start + 1, start_column=1, end_row=start + 12, end_column=3)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# =============================================================================
# API endpoints
# =============================================================================

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "app": "SteelDraw AI Extractor",
        "openai_configured": bool(OPENAI_API_KEY),
    }


@app.post("/api/upload")
async def upload(
    file: Optional[UploadFile] = File(None),
    plan_pages: Optional[str] = Form(None),
    elevation_pages: Optional[str] = Form(None),
    job_id: Optional[str] = Form(None),
    confirm: Optional[str] = Form(None),
):
    """
    POST /api/upload

    Flow A — first upload (no page selection yet):
      - Saves PDF, returns page_count.
      - If page_count > 5 and confirm != 'true': returns needs_page_selection=true
        so the UI can ask for plan vs elevation pages.
      - If page_count <= 5: processes all pages and returns Excel + JSON preview.

    Flow B — follow-up with job_id + plan_pages + elevation_pages + confirm=true:
      - Processes selected pages and returns Excel + JSON preview.

    Excel is returned as base64 in JSON along with preview tables so the React
    UI can show results and offer Download without a second round-trip.
    """
    import base64

    # --- Resolve PDF path (new upload or existing job) ---
    pdf_path: Optional[str] = None
    filename = "drawing.pdf"
    page_count = 0

    if job_id and job_id in JOBS:
        job = JOBS[job_id]
        pdf_path = job["pdf_path"]
        filename = job["filename"]
        page_count = job["page_count"]
        if not Path(pdf_path).exists():
            raise HTTPException(status_code=400, detail="Job expired — please re-upload the PDF.")
    else:
        if file is None or not file.filename or not file.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail="Please upload a PDF file.")

        # Stream to disk to support up to 500MB without loading all into RAM
        suffix = Path(file.filename).suffix or ".pdf"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        pdf_path = tmp.name
        filename = file.filename
        total = 0
        try:
            while True:
                chunk = await file.read(1024 * 1024)  # 1 MB chunks
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    tmp.close()
                    os.unlink(pdf_path)
                    raise HTTPException(status_code=413, detail="File exceeds 500 MB limit.")
                tmp.write(chunk)
            tmp.close()
        except HTTPException:
            raise
        except Exception as exc:  # noqa: BLE001
            tmp.close()
            if Path(pdf_path).exists():
                os.unlink(pdf_path)
            raise HTTPException(status_code=400, detail=f"Failed to save upload: {exc}") from exc

        try:
            page_count = get_page_count(pdf_path)
        except Exception as exc:  # noqa: BLE001
            os.unlink(pdf_path)
            raise HTTPException(status_code=400, detail=f"Invalid PDF: {exc}") from exc

        new_job_id = str(uuid.uuid4())
        JOBS[new_job_id] = {
            "pdf_path": pdf_path,
            "page_count": page_count,
            "filename": filename,
        }
        job_id = new_job_id

    # --- Page selection gate for large drawings ---
    confirmed = (confirm or "").lower() in ("1", "true", "yes")
    if page_count > PAGE_ASK_THRESHOLD and not confirmed:
        return JSONResponse(
            {
                "needs_page_selection": True,
                "job_id": job_id,
                "page_count": page_count,
                "filename": filename,
                "message": (
                    f"This drawing has {page_count} pages. "
                    "Select which plan page(s) to use for beam & bracing details, "
                    "and which elevation page(s) to use for column & elevation / base plate details."
                ),
            }
        )

    # --- Determine pages to scan ---
    if page_count <= PAGE_ASK_THRESHOLD:
        plan = list(range(1, page_count + 1))
        elev = list(range(1, page_count + 1))
    else:
        plan = parse_page_list(plan_pages, page_count)
        elev = parse_page_list(elevation_pages, page_count)
        if not plan and not elev:
            raise HTTPException(
                status_code=400,
                detail="Provide plan_pages and/or elevation_pages (e.g. '1,2' or '3-5').",
            )

    try:
        result = process_pdf(pdf_path, plan_pages=plan, elevation_pages=elev)
        excel_bytes = build_excel(result)
    except HTTPException:
        _cleanup_job = confirmed or page_count <= PAGE_ASK_THRESHOLD
        if _cleanup_job:
            job = JOBS.pop(job_id, None)
            if job and Path(job["pdf_path"]).exists():
                try:
                    os.unlink(job["pdf_path"])
                except OSError:
                    pass
        raise
    except Exception as exc:  # noqa: BLE001
        _cleanup_job = confirmed or page_count <= PAGE_ASK_THRESHOLD
        if _cleanup_job:
            job = JOBS.pop(job_id, None)
            if job and Path(job["pdf_path"]).exists():
                try:
                    os.unlink(job["pdf_path"])
                except OSError:
                    pass
        raise HTTPException(status_code=500, detail=f"Extraction failed: {exc}") from exc

    # Cleanup after successful extract
    if confirmed or page_count <= PAGE_ASK_THRESHOLD:
        job = JOBS.pop(job_id, None)
        if job and Path(job["pdf_path"]).exists():
            try:
                os.unlink(job["pdf_path"])
            except OSError:
                pass

    b64 = base64.b64encode(excel_bytes).decode("ascii")
    download_name = f"{Path(filename).stem}_SteelDraw_Extract.xlsx"

    return JSONResponse(
        {
            "needs_page_selection": False,
            "job_id": job_id,
            "page_count": page_count,
            "filename": filename,
            "download_filename": download_name,
            "excel_base64": b64,
            "preview": {
                "beams": result["beams"],
                "columns": result["columns"],
                "base_plates": result["base_plates"],
                "bracings": result["bracings"],
                "summary": result["summary"],
                "engineer_notes": result["engineer_notes"],
            },
            "counts": {
                "beams": len(result["beams"]),
                "columns": len(result["columns"]),
                "base_plates": len(result["base_plates"]),
                "bracings": len(result["bracings"]),
            },
            "page_sources": result["page_sources"],
        }
    )


@app.post("/api/download")
async def download_excel_only(
    file: UploadFile = File(...),
    plan_pages: Optional[str] = Form(None),
    elevation_pages: Optional[str] = Form(None),
):
    """
    Alternate endpoint: process and stream Excel directly as a file download.
    Useful for API clients; the React UI uses /api/upload for preview + download.
    """
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf_path = tmp.name
    try:
        total = 0
        while True:
            chunk = await file.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="File exceeds 500 MB limit.")
            tmp.write(chunk)
        tmp.close()

        page_count = get_page_count(pdf_path)
        if page_count > PAGE_ASK_THRESHOLD and not (plan_pages or elevation_pages):
            raise HTTPException(
                status_code=400,
                detail=(
                    f"PDF has {page_count} pages. Pass plan_pages and elevation_pages "
                    "form fields (e.g. plan_pages=1,2&elevation_pages=3-5)."
                ),
            )
        plan = parse_page_list(plan_pages, page_count)
        elev = parse_page_list(elevation_pages, page_count)
        result = process_pdf(pdf_path, plan_pages=plan, elevation_pages=elev)
        excel_bytes = build_excel(result)
    finally:
        tmp.close()
        if Path(pdf_path).exists():
            os.unlink(pdf_path)

    name = f"{Path(file.filename).stem}_SteelDraw_Extract.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
