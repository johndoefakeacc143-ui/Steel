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
_ROOT_ENV = Path(__file__).resolve().parent.parent / ".env"
_BACKEND_ENV = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=_ROOT_ENV)
load_dotenv(dotenv_path=_BACKEND_ENV)
load_dotenv()  # also allow process env


def _refresh_openai_key() -> str:
    """
    Re-read .env on each AI call so editing OPENAI_API_KEY takes effect
    after save (uvicorn --reload also picks it up on file change).
    """
    load_dotenv(dotenv_path=_ROOT_ENV, override=True)
    load_dotenv(dotenv_path=_BACKEND_ENV, override=True)
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    # Treat placeholder values as empty
    if not key or key.lower().startswith("sk-your") or key == "sk-...":
        return ""
    return key


OPENAI_API_KEY = _refresh_openai_key()
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
# Negative lookbehind/ahead so BR1 / BP2 are NOT treated as beam marks.
BEAM_MARK_RE = re.compile(
    r"(?<![A-Z])(?:B|BM|BEAM)[-\s]?(\d{1,4}[A-Z]?)\b",
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

# Standalone plan dimension values (e.g. 1500, 2000, 3000, 6000).
# Edit DIM_VALUE_RE if your sheets use different ranges / units.
DIM_VALUE_RE = re.compile(
    r"^(?P<val>\d{3,5}(?:\.\d+)?)(?:\s*(?:mm|cm|m))?$",
    re.IGNORECASE,
)

# Explicit bay / span pairs sometimes written as 3000x2000 near a brace
BAY_PAIR_RE = re.compile(
    r"\b(\d{3,5}(?:\.\d+)?)\s*[x×]\s*(\d{3,5}(?:\.\d+)?)\b",
    re.IGNORECASE,
)


def _normalize_section(raw: str) -> str:
    return re.sub(r"\s+", "", raw.upper().replace("×", "x"))


def _parse_length_mm(match: re.Match) -> Optional[float]:
    """Convert LENGTH_RE match groups to millimetres."""
    if match.group(1):
        return float(match.group(1))
    if match.group(2):
        return float(match.group(2)) * 1000.0
    return None


def _to_mm(value: float, unit_hint: str = "") -> float:
    """
    Normalize a raw dimension to millimetres.
    Plan grids like 6000 / 1500 / 2000 are almost always mm on steel drawings.
    Only convert when the unit is explicitly written.
    """
    u = (unit_hint or "").strip().lower()
    if u in ("m", "meter", "metre", "meters", "metres"):
        return value * 1000.0
    if u in ("cm", "centimeter", "centimetre"):
        return value * 10.0
    return value  # default mm


def triangle_diagonal_length(a: float, b: float) -> float:
    """
    Diagonal / bracing length from bay rectangle sides (Pythagoras):
        L = √(a² + b²)
    Example: bay 3000 × 2000 → √(3000² + 2000²) = 3605.55
    """
    return round(float(np.sqrt(a * a + b * b)), 2)


def _nearby_window(text: str, start: int, end: int, radius: int = 180) -> str:
    """
    Prefer the mark's own line plus at most one adjacent line.
    Avoid huge character windows on compact plan OCR dumps where every mark
    would otherwise see the whole page (and pick up unrelated 'diagonal' notes).
    """
    line_start = text.rfind("\n", 0, start) + 1
    line_end = text.find("\n", end)
    if line_end == -1:
        line_end = len(text)
    line = text[line_start:line_end].strip()

    # Collect previous / next line for wrapped schedule notes
    prev_start = text.rfind("\n", 0, max(0, line_start - 1)) + 1
    prev_line = text[prev_start:line_start].strip() if line_start > 0 else ""
    next_end = text.find("\n", line_end + 1)
    if next_end == -1:
        next_end = min(len(text), line_end + 80)
    next_line = text[line_end:next_end].strip()

    parts = [line]
    # Only attach neighbour lines that do NOT introduce another member mark
    def _is_other_mark_line(s: str) -> bool:
        if not s:
            return False
        return bool(
            BEAM_MARK_RE.search(s)
            or COLUMN_MARK_RE.search(s)
            or BASE_PLATE_MARK_RE.search(s)
            or BRACING_MARK_RE.search(s)
        )

    if next_line and not _is_other_mark_line(next_line):
        parts.append(next_line)
    elif next_line and len(next_line) <= 12 and next_line.replace(".", "").isdigit():
        # Bare dimension on the next line (common under a mark)
        parts.append(next_line)

    if prev_line and not _is_other_mark_line(prev_line):
        if len(prev_line) <= 12 and prev_line.replace(".", "").isdigit():
            parts.insert(0, prev_line)

    window = " ".join(p for p in parts if p)
    if len(window) >= 2:
        return window

    lo = max(0, start - min(radius, 40))
    hi = min(len(text), end + min(radius, 40))
    return text[lo:hi]


def _mark_local_text(text: str, mark: str, radius: int = 80) -> str:
    """
    Return text near a whole-word mark only (so B4 does not hit inside BR4).
    """
    if not text or not mark:
        return ""
    pattern = re.compile(rf"(?<![A-Z0-9]){re.escape(mark)}(?![A-Z0-9])", re.IGNORECASE)
    m = pattern.search(text)
    if not m:
        return ""
    return _nearby_window(text, m.start(), m.end(), radius=radius)


def length_from_parallel_dimension(
    mark_item: dict[str, Any],
    dims: list[dict[str, Any]],
    preferred_direction: Optional[str] = None,
) -> tuple[Optional[float], str]:
    """
    Pick the dimension written in the SAME DIRECTION as the member.

    On plans:
      - Vertical beam (B3, B7, B8) → vertical dim beside it (1500 / 2000 / 6000)
      - Horizontal beam (B4) → horizontal dim above/below (6000)

    Among well-aligned dims, choose the NEAREST one (so B3 gets 1500, not a
    far-away 6000 that happens to share the same X grid line).
    """
    if not dims:
        return None, ""

    direction = preferred_direction or mark_item.get("orientation") or "unknown"
    mx, my = mark_item["cx"], mark_item["cy"]

    def collect(direction_name: str) -> list[tuple[float, float, dict[str, Any]]]:
        """Return list of (align_err, dist, dim) for viable candidates."""
        out = []
        for d in dims:
            dx = abs(d["cx"] - mx)
            dy = abs(d["cy"] - my)
            dist = float(np.hypot(dx, dy))
            if dist < 8:
                continue
            if direction_name == "vertical":
                align, along = dx, dy
                orient_ok = d.get("orientation") in ("vertical", "unknown")
            else:
                align, along = dy, dx
                orient_ok = d.get("orientation") in ("horizontal", "unknown")
            # Must be offset along the member, and tightly aligned perpendicular
            if align > 80:
                continue
            if along < 12:
                continue
            # Soft cap: don't jump across the whole sheet for a "same grid line" dim
            if along > 220 and align > 25:
                continue
            if along > 350:
                continue
            # Prefer matching orientation; still allow unknown
            if not orient_ok and d.get("orientation") not in ("unknown", None):
                # opposite orientation — skip for this direction pass
                continue
            out.append((align, dist, d))
        return out

    directions = (
        [direction] if direction in ("vertical", "horizontal") else ["vertical", "horizontal"]
    )

    best: Optional[tuple[float, float, dict, str]] = None
    for dir_name in directions:
        cands = collect(dir_name)
        if not cands:
            continue
        # Sort by distance first, then alignment
        cands.sort(key=lambda t: (t[1], t[0]))
        align, dist, dim = cands[0]
        # Bonus: if dim orientation matches dir, prefer it over opposite-dir result
        score_key = (dist, align)
        if best is None or score_key < (best[0], best[1]):
            best = (dist, align, dim, dir_name)

    if best is None:
        nearest = min(dims, key=lambda d: float(np.hypot(d["cx"] - mx, d["cy"] - my)))
        return nearest.get("value_mm"), "nearest"

    return best[2].get("value_mm"), best[3]


def _first_section(window: str) -> str:
    m = SECTION_SIZE_RE.search(window)
    return _normalize_section(m.group(0)) if m else ""


def _first_length(window: str) -> Optional[float]:
    # Prefer explicit L= / Length / Height patterns first
    # Ignore survey/grid coordinates like N=3375.000 / E=4089.500 — not member lengths
    cleaned = re.sub(
        r"\b[NE]\s*=\s*[+\-]?\d+(?:\.\d+)?",
        " ",
        window,
        flags=re.IGNORECASE,
    )
    explicit = re.search(
        r"(?:L|LEN(?:GTH)?|H(?:EIGHT)?|HT)\s*[=:]?\s*(\d{3,5}(?:\.\d+)?)\s*(?:mm)?"
        r"|(?:L|LEN(?:GTH)?|H(?:EIGHT)?|HT)\s*[=:]?\s*(\d{1,2}(?:\.\d+)?)\s*m\b",
        cleaned,
        re.IGNORECASE,
    )
    if explicit:
        if explicit.group(1):
            return float(explicit.group(1))
        if explicit.group(2):
            return float(explicit.group(2)) * 1000.0
    for m in LENGTH_RE.finditer(cleaned):
        val = _parse_length_mm(m)
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


def _word_center(w: dict[str, Any]) -> tuple[float, float]:
    return ((w["x0"] + w["x1"]) / 2.0, (w["top"] + w["bottom"]) / 2.0)


def _word_orientation(w: dict[str, Any]) -> str:
    """Infer text direction from bounding-box aspect ratio."""
    width = max(0.1, w["x1"] - w["x0"])
    height = max(0.1, w["bottom"] - w["top"])
    if height > width * 1.25:
        return "vertical"
    if width > height * 1.25:
        return "horizontal"
    return "unknown"


def _normalize_dim_token(raw: str, orientation: str = "unknown") -> Optional[float]:
    """
    Parse a dimension token to mm.
    Rotated vertical dims are often extracted reversed (1500 → '0051').
    Try the token as-is, then reversed when orientation is vertical.
    """
    candidates = [raw]
    if orientation == "vertical" and raw.isdigit() and len(raw) >= 3:
        candidates.append(raw[::-1])
    # Also try reverse when the forward value looks like a leading-zero artifact
    if raw.isdigit() and raw.startswith("0") and len(raw) >= 3:
        candidates.append(raw[::-1])

    for cand in candidates:
        dm = DIM_VALUE_RE.fullmatch(cand)
        if not dm:
            # bare digits without unit
            if re.fullmatch(r"\d{3,5}(?:\.\d+)?", cand):
                raw_val = float(cand)
            else:
                continue
        else:
            raw_val = float(dm.group("val"))
        if 100 <= raw_val <= 30000:
            # Prefer values that don't look like reversed leftovers starting with 0
            if cand.startswith("0") and len(cand) >= 4 and orientation == "vertical":
                continue
            unit = ""
            um = re.search(r"(mm|cm|m)$", cand, re.IGNORECASE)
            if um:
                unit = um.group(1)
            return _to_mm(raw_val, unit)
    return None


def collect_page_geometry(page: pdfplumber.page.Page) -> dict[str, list[dict[str, Any]]]:
    """
    Pull words with coordinates from a digital PDF page.
    Returns marks (beams/columns/bracings/baseplates) and dimension values.

    Also rebuilds vertically-stacked / rotated dimension numbers from raw
    character positions (common on plan drawings where dims run along the beam).
    """
    try:
        words = page.extract_words(
            use_text_flow=False,
            keep_blank_chars=False,
            extra_attrs=["size"],
        ) or []
    except Exception as exc:  # noqa: BLE001
        print(f"[geometry] extract_words failed: {exc}")
        words = []

    beams: list[dict[str, Any]] = []
    columns: list[dict[str, Any]] = []
    bracings: list[dict[str, Any]] = []
    base_plates: list[dict[str, Any]] = []
    dims: list[dict[str, Any]] = []

    def _add_mark(raw: str, item: dict[str, Any]) -> bool:
        if re.fullmatch(r"(?:B|BM|BEAM)[-\s]?\d{1,4}[A-Z]?", raw, re.IGNORECASE):
            m = BEAM_MARK_RE.search(raw)
            if m:
                # Skip absurd marks from drawing numbers (B402 from 402-R-...)
                num = m.group(1)
                if num.isdigit() and int(num) > 200:
                    return False
                item["mark"] = f"B{num.upper()}"
                beams.append(item)
                return True
        if re.fullmatch(r"(?:C|COL|COLUMN)[-\s]?\d{1,4}[A-Z]?", raw, re.IGNORECASE):
            m = COLUMN_MARK_RE.search(raw)
            if m:
                num = m.group(1)
                if num.isdigit() and int(num) > 200:
                    return False
                item["mark"] = f"C{num.upper()}"
                columns.append(item)
                return True
        if re.fullmatch(r"(?:BR|BRG|BRACING)[-\s]?\d{1,4}[A-Z]?", raw, re.IGNORECASE):
            m = BRACING_MARK_RE.search(raw)
            if m:
                item["mark"] = f"BR{m.group(1).upper()}"
                bracings.append(item)
                return True
        if re.fullmatch(r"(?:BP|B\.?P\.?)[-\s]?\d{1,4}[A-Z]?", raw, re.IGNORECASE):
            m = BASE_PLATE_MARK_RE.search(raw)
            if m:
                item["mark"] = f"BP{m.group(1).upper()}"
                base_plates.append(item)
                return True
        return False

    for w in words:
        raw = (w.get("text") or "").strip()
        if not raw:
            continue
        cx, cy = _word_center(w)
        orient = _word_orientation(w)
        # Rotated vertical text often has taller bbox
        width = max(0.1, w["x1"] - w["x0"])
        height = max(0.1, w["bottom"] - w["top"])
        if height > width * 1.2:
            orient = "vertical"
        item = {
            "text": raw,
            "x0": w["x0"],
            "x1": w["x1"],
            "top": w["top"],
            "bottom": w["bottom"],
            "cx": cx,
            "cy": cy,
            "orientation": orient,
        }
        if _add_mark(raw, item):
            continue

        # Dimension values (including reversed rotated tokens like 0051 → 1500)
        val = _normalize_dim_token(raw, orient)
        if val is not None:
            item["value_mm"] = val
            item["raw_value"] = val
            dims.append(item)

    # --- Rebuild rotated dims from chars when upright=False ---
    try:
        chars = page.chars or []
    except Exception:  # noqa: BLE001
        chars = []

    rotated_digits = [
        ch for ch in chars
        if (ch.get("text") or "").strip().isdigit() and ch.get("upright") is False
    ]
    if rotated_digits:
        rotated_digits.sort(key=lambda c: ((c["x0"] + c["x1"]) / 2.0, -c["top"]))
        clusters: list[list] = []
        for ch in rotated_digits:
            cx = (ch["x0"] + ch["x1"]) / 2.0
            placed = False
            for cluster in clusters:
                ccx = sum((c["x0"] + c["x1"]) / 2.0 for c in cluster) / len(cluster)
                if abs(cx - ccx) <= 6.0:
                    cluster.append(ch)
                    placed = True
                    break
            if not placed:
                clusters.append([ch])

        existing = {round(d["value_mm"], 2) for d in dims}
        for cluster in clusters:
            if len(cluster) < 3:
                continue
            # For rotate(90) / upright=False: sort by top descending → correct reading
            cluster.sort(key=lambda c: -c["top"])
            text_val = "".join(c["text"] for c in cluster)
            if re.fullmatch(r"\d{3,5}", text_val):
                raw_val = float(text_val)
                if 100 <= raw_val <= 30000:
                    cx = sum((c["x0"] + c["x1"]) / 2.0 for c in cluster) / len(cluster)
                    cy = (min(c["top"] for c in cluster) + max(c["bottom"] for c in cluster)) / 2.0
                    # Avoid dupes already recovered via reversed words
                    if any(
                        abs(d["value_mm"] - raw_val) < 0.1 and abs(d["cx"] - cx) < 15
                        for d in dims
                    ):
                        continue
                    dims.append(
                        {
                            "text": text_val,
                            "x0": min(c["x0"] for c in cluster),
                            "x1": max(c["x1"] for c in cluster),
                            "top": min(c["top"] for c in cluster),
                            "bottom": max(c["bottom"] for c in cluster),
                            "cx": cx,
                            "cy": cy,
                            "orientation": "vertical",
                            "value_mm": raw_val,
                            "raw_value": raw_val,
                        }
                    )
                    existing.add(raw_val)

    return {
        "beams": beams,
        "columns": columns,
        "bracings": bracings,
        "base_plates": base_plates,
        "dims": dims,
    }


def ocr_plan_dimensions(page: pdfplumber.page.Page, dpi: int = 120) -> list[dict[str, Any]]:
    """
    Harvest dimension callouts that live in the drawing graphics (not the text
    layer). Many steel plans only embed member marks as text; sizes like 1500 /
    2000 / 6000 are drawn as graphics and need OCR.

    Runs upright + 90° rotations so vertical dimension strings are found.
    Coordinates are mapped back into pdfplumber page space.
    """
    try:
        pil = page.to_image(resolution=dpi).original
    except Exception as exc:  # noqa: BLE001
        print(f"[OCR dims] render failed: {exc}")
        return []

    arr = np.array(pil)
    page_w, page_h = float(page.width), float(page.height)
    img_h, img_w = arr.shape[:2]
    sx = page_w / img_w
    sy = page_h / img_h

    dim_re = re.compile(r"^\d{3,5}$")
    found: list[dict[str, Any]] = []

    def _ingest(gray: np.ndarray, transform: str) -> None:
        try:
            data = pytesseract.image_to_data(
                gray,
                config="--psm 11 -c tessedit_char_whitelist=0123456789",
                output_type=pytesseract.Output.DICT,
            )
        except Exception as exc:  # noqa: BLE001
            print(f"[OCR dims] tesseract failed ({transform}): {exc}")
            return
        n = len(data.get("text") or [])
        for i in range(n):
            raw = (data["text"][i] or "").strip()
            if not dim_re.fullmatch(raw):
                continue
            try:
                conf = float(data["conf"][i])
            except ValueError:
                conf = -1.0
            if conf < 45:
                continue
            val = float(raw)
            if val < 400 or val > 20000:
                continue
            left, top = int(data["left"][i]), int(data["top"][i])
            width, height = int(data["width"][i]), int(data["height"][i])
            if transform == "0":
                ix, iy = left + width / 2.0, top + height / 2.0
                orient = "horizontal" if width >= height else "vertical"
            elif transform == "90cw":
                ix = top + height / 2.0
                iy = img_h - 1 - (left + width / 2.0)
                orient = "vertical"
            elif transform == "90ccw":
                ix = img_w - 1 - (top + height / 2.0)
                iy = left + width / 2.0
                orient = "vertical"
            else:
                continue
            px, py = ix * sx, iy * sy
            found.append(
                {
                    "text": raw,
                    "x0": px - 5,
                    "x1": px + 5,
                    "top": py - 5,
                    "bottom": py + 5,
                    "cx": px,
                    "cy": py,
                    "orientation": orient,
                    "value_mm": val,
                    "raw_value": val,
                    "source": f"ocr-{transform}",
                }
            )

    gray0 = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    _ingest(gray0, "0")
    _ingest(cv2.rotate(gray0, cv2.ROTATE_90_CLOCKWISE), "90cw")
    _ingest(cv2.rotate(gray0, cv2.ROTATE_90_COUNTERCLOCKWISE), "90ccw")

    unique: list[dict[str, Any]] = []
    for d in found:
        if any(
            abs(d["value_mm"] - u["value_mm"]) < 1
            and abs(d["cx"] - u["cx"]) < 20
            and abs(d["cy"] - u["cy"]) < 20
            for u in unique
        ):
            continue
        unique.append(d)
    print(f"[OCR dims] harvested {len(unique)} dimension callouts")
    return unique


def estimate_scale_mm_per_unit(page: pdfplumber.page.Page) -> Optional[float]:
    """
    Estimate mm-per-page-unit from PB1–PB2 grid labels assumed 6000 mm.
    Edit GRID_SPAN_MM below if your primary grid is not 6000.
    """
    GRID_SPAN_MM = 6000.0  # <-- edit if primary grid is not 6000
    try:
        words = page.extract_words(use_text_flow=False) or []
    except Exception:  # noqa: BLE001
        return None
    pb1, pb2 = [], []
    for w in words:
        t = (w.get("text") or "").strip().upper()
        cx = (w["x0"] + w["x1"]) / 2.0
        if t == "PB1":
            pb1.append(cx)
        elif t == "PB2":
            pb2.append(cx)
    if not pb1 or not pb2:
        return None
    spans = []
    for a in pb1:
        rights = [b for b in pb2 if b > a + 40]
        if rights:
            spans.append(min(rights) - a)
    if not spans:
        return None
    avg = sum(spans) / len(spans)
    if avg < 20:
        return None
    return GRID_SPAN_MM / avg


def length_from_mark_spacing(
    mark_item: dict[str, Any],
    same_mark_points: list[dict[str, Any]],
    scale: float,
    preferred_direction: Optional[str] = None,
) -> tuple[Optional[float], str]:
    """
    Fallback when OCR dims are sparse: estimate length from spacing to the
    nearest same-mark neighbour along the member direction, scaled by grid.
    """
    if scale <= 0 or not same_mark_points:
        return None, ""
    mx, my = mark_item["cx"], mark_item["cy"]
    direction = preferred_direction or "unknown"
    best = None
    for p in same_mark_points:
        if abs(p["cx"] - mx) < 0.5 and abs(p["cy"] - my) < 0.5:
            continue
        dx, dy = abs(p["cx"] - mx), abs(p["cy"] - my)
        dist = float(np.hypot(dx, dy))
        if dist < 15:
            continue
        if direction == "vertical" or (direction == "unknown" and dy >= dx):
            if dx > 40:
                continue
            length = round(dy * scale)
            how = "spacing-vertical"
        else:
            if dy > 40:
                continue
            length = round(dx * scale)
            how = "spacing-horizontal"
        if 400 <= length <= 20000:
            if best is None or dist < best[0]:
                best = (dist, float(length), how)
    if best:
        return best[1], best[2]
    return None, ""


def diagonal_length_from_bay(
    brace_item: dict[str, Any],
    dims: list[dict[str, Any]],
    text_window: str = "",
) -> tuple[Optional[float], str]:
    """
    For diagonally placed members (bracings / sloping beams):
      L = √(a² + b²)

    a = nearest horizontal bay dimension, b = nearest vertical bay dimension.
    Also accepts explicit pairs in nearby text: '3000x2000'.
    """
    # 1) Explicit bay pair in text near the mark
    pair = BAY_PAIR_RE.search(text_window or "")
    if pair:
        a, b = float(pair.group(1)), float(pair.group(2))
        # Ignore plate-like triples already handled elsewhere (e.g. 600x600)
        if a >= 200 and b >= 200:
            return triangle_diagonal_length(a, b), f"√({a}²+{b}²) from text pair"

    if not dims:
        return None, ""

    mx, my = brace_item["cx"], brace_item["cy"]

    # Split dims into likely horizontal-string vs vertical-string values
    horiz: list[tuple[float, dict]] = []
    vert: list[tuple[float, dict]] = []
    for d in dims:
        dx = abs(d["cx"] - mx)
        dy = abs(d["cy"] - my)
        dist = float(np.hypot(dx, dy))
        if dist < 10 or dist > 500:
            continue
        # Horizontal dimension strings sit above/below the bay (similar Y band offset)
        # Vertical dimension strings sit left/right (similar X band offset)
        h_score = dist + dx * 0.5  # prefer closer in Y for horizontal dims
        v_score = dist + dy * 0.5
        if d.get("orientation") == "horizontal":
            h_score -= 30
        if d.get("orientation") == "vertical":
            v_score -= 30
        # Dims nearly aligned horizontally with brace center → vertical dim line
        if dx < 90:
            vert.append((v_score, d))
        if dy < 90:
            horiz.append((h_score, d))
        # Also keep general nearest pools
        if dy <= dx:
            horiz.append((h_score + 20, d))
        else:
            vert.append((v_score + 20, d))

    if not horiz or not vert:
        # Fallback: two nearest distinct dim values as a,b
        ordered = sorted(dims, key=lambda d: float(np.hypot(d["cx"] - mx, d["cy"] - my)))
        vals = []
        for d in ordered:
            v = d.get("value_mm")
            if v and v not in vals:
                vals.append(v)
            if len(vals) >= 2:
                break
        if len(vals) >= 2:
            a, b = vals[0], vals[1]
            return triangle_diagonal_length(a, b), f"√({a}²+{b}²) nearest dims"
        return None, ""

    horiz.sort(key=lambda t: t[0])
    vert.sort(key=lambda t: t[0])
    a = horiz[0][1]["value_mm"]
    b = vert[0][1]["value_mm"]
    # Avoid using the same physical dim twice when pools overlap
    if a == b and len(horiz) > 1:
        a = horiz[1][1]["value_mm"]
    if a == b and len(vert) > 1:
        b = vert[1][1]["value_mm"]
    return triangle_diagonal_length(a, b), f"√({a}²+{b}²)"



def apply_geometry_lengths(
    rows: list[dict[str, Any]],
    geo_marks: list[dict[str, Any]],
    dims: list[dict[str, Any]],
    length_key: str,
    diagonal: bool = False,
    text: str = "",
    scale: Optional[float] = None,
) -> list[dict[str, Any]]:
    """
    Attach / override lengths using plan geometry.
    Keeps one row per mark occurrence when the same mark appears multiple times
    (e.g. two B8 members both length 6000).
    """
    if not geo_marks:
        return rows

    # Index existing text-extracted rows by mark for section/material reuse
    by_mark: dict[str, dict[str, Any]] = {}
    for r in rows:
        by_mark.setdefault(r.get("Mark") or "", r)

    # Group geometry marks by mark id for spacing fallback
    by_geo_mark: dict[str, list[dict[str, Any]]] = {}
    for g in geo_marks:
        by_geo_mark.setdefault(g.get("mark") or "", []).append(g)

    built: list[dict[str, Any]] = []
    for idx, g in enumerate(geo_marks):
        mark = g.get("mark") or ""
        base = dict(by_mark.get(mark) or {"Mark": mark})
        base["Mark"] = mark
        # Stable instance id so two B8@6000 rows are not collapsed later
        base["_instance"] = f"{mark}-{idx}-{round(g.get('cx', 0))}-{round(g.get('cy', 0))}"

        if diagonal:
            local = _mark_local_text(text, mark, radius=160)
            length, how = diagonal_length_from_bay(g, dims, local)
            if length is None and scale:
                # Bay fallback: use nearest horiz+vert OCR/spacing dims already in diagonal_length_from_bay
                pass
            if length is not None:
                base[length_key] = length
                base["Length Method"] = how
        else:
            preferred = g.get("orientation")
            if preferred not in ("vertical", "horizontal"):
                preferred = None
            length, direction = length_from_parallel_dimension(
                g, dims, preferred_direction=preferred
            )
            # Prefer grid-spacing when parallel association fell back to "nearest"
            # (weak match) or found nothing — spacing is reliable on regular racks.
            weak = length is None or direction == "nearest"
            if weak and scale:
                s_len, s_how = length_from_mark_spacing(
                    g,
                    by_geo_mark.get(mark) or [],
                    scale,
                    preferred_direction=preferred,
                )
                if s_len is not None:
                    # If we had a weak nearest dim, only override when spacing
                    # is a clean structural length (multiples of 500).
                    if length is None or (
                        direction == "nearest" and s_len % 500 == 0
                    ):
                        length, direction = s_len, s_how

            existing = base.get(length_key)
            # Prefer explicit schedule lengths (L=3000) over nearby grid dims.
            has_explicit = existing not in ("", None) and not base.get("Length Method")
            if length is not None and not has_explicit:
                base[length_key] = length
                if str(direction).startswith("spacing-"):
                    base["Length Method"] = direction
                    base["Member Direction"] = direction.replace("spacing-", "")
                else:
                    base["Length Method"] = f"parallel-{direction}"
                    if direction in ("vertical", "horizontal"):
                        base["Member Direction"] = direction
            elif has_explicit:
                base["Length Method"] = base.get("Length Method") or "explicit-text"
                if length is not None and direction in ("vertical", "horizontal"):
                    base["Member Direction"] = direction

        built.append(base)

    if not built:
        return rows

    geo_mark_set = {g.get("mark") for g in geo_marks}
    extras = [r for r in rows if (r.get("Mark") or "") not in geo_mark_set]
    return built + extras


def extract_beams_regex(
    text: str,
    page_no: int,
    geometry: Optional[dict[str, list]] = None,
    scale: Optional[float] = None,
) -> list[dict[str, Any]]:
    """
    Extract beam rows.
    Length rule (plan drawings): use the dimension written in the SAME DIRECTION
    as the beam (vertical dim for vertical beams, horizontal dim for horizontal).
    Example from typical plans: B3=1500, B8=6000, B7=2000, B4=6000.
    """
    rows: list[dict[str, Any]] = []
    seen_spans: set[tuple[int, int]] = set()
    for m in BEAM_MARK_RE.finditer(text):
        # Guard: character before match must not make it part of BR/BP
        if m.start() > 0 and text[m.start() - 1].upper() in ("R", "P"):
            continue
        num = m.group(1)
        if num.isdigit() and int(num) > 200:
            continue
        mark = f"B{num.upper()}"
        span = (m.start(), m.end())
        if span in seen_spans:
            continue
        seen_spans.add(span)
        window = _nearby_window(text, m.start(), m.end())
        elevs = _elevations(window)
        length = _first_length(window)
        # Ignore lengths that clearly came from "3000x2000" bay pairs unless
        # the beam is marked diagonal (handled later).
        if BAY_PAIR_RE.search(window) and not re.search(
            r"\b(DIAGONAL|SLOPING|INCLINED|SLOPED)\b", window, re.IGNORECASE
        ):
            explicit = re.search(
                r"(?:L|LEN(?:GTH)?)\s*[=:]\s*(\d{3,5}(?:\.\d+)?)",
                window,
                re.IGNORECASE,
            )
            length = float(explicit.group(1)) if explicit else None

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

    # Deduplicate identical text marks before geometry expands instances
    rows = dedupe_by_mark(rows, allow_multi=False)

    if geometry:
        rows = apply_geometry_lengths(
            rows,
            geometry.get("beams") or [],
            geometry.get("dims") or [],
            length_key="Length (mm)",
            diagonal=False,
            text=text,
            scale=scale,
        )
        for r in rows:
            r.setdefault("Page", page_no)
    return rows


def extract_columns_regex(
    text: str,
    page_no: int,
    geometry: Optional[dict[str, list]] = None,
    scale: Optional[float] = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in COLUMN_MARK_RE.finditer(text):
        num = m.group(1)
        if num.isdigit() and int(num) > 200:
            continue
        mark = f"C{num.upper()}"
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

    # On ELEVATION views columns are vertical — associate vertical dims when present
    if geometry:
        rows = apply_geometry_lengths(
            rows,
            geometry.get("columns") or [],
            geometry.get("dims") or [],
            length_key="Height (mm)",
            diagonal=False,
            text=text,
            scale=scale,
        )
        for r in rows:
            r.setdefault("Page", page_no)
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


def extract_bracings_regex(
    text: str,
    page_no: int,
    geometry: Optional[dict[str, list]] = None,
    scale: Optional[float] = None,
) -> list[dict[str, Any]]:
    """
    Bracings (and any diagonally placed member) — length from triangle diagonal:
        L = √(a² + b²)
    where a, b are the horizontal and vertical bay dimensions the brace spans.
    Example: bay 3000 × 2000 → L = 3605.55
    """
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for m in BRACING_MARK_RE.finditer(text):
        mark = f"BR{m.group(1).upper()}"
        if mark in seen:
            continue
        seen.add(mark)
        window = _nearby_window(text, m.start(), m.end(), radius=160)
        length = None
        method = ""
        pair = BAY_PAIR_RE.search(window)
        if pair:
            a, b = float(pair.group(1)), float(pair.group(2))
            if a >= 200 and b >= 200:
                length = triangle_diagonal_length(a, b)
                method = f"√({a}²+{b}²) from text pair"
        if length is None:
            # Explicit L= hypotenuse already written on the sheet
            explicit = re.search(
                r"(?:L|LEN(?:GTH)?)\s*[=:]\s*(\d{3,5}(?:\.\d+)?)",
                window,
                re.IGNORECASE,
            )
            if explicit:
                length = float(explicit.group(1))
                method = "explicit"
        rows.append(
            {
                "Mark": mark,
                "Section Size": _first_section(window),
                "Length (mm)": length if length is not None else "",
                "Length Method": method,
                "Page": page_no,
            }
        )

    if geometry:
        rows = apply_geometry_lengths(
            rows,
            geometry.get("bracings") or [],
            geometry.get("dims") or [],
            length_key="Length (mm)",
            diagonal=True,
            text=text,
            scale=scale,
        )
        for r in rows:
            r.setdefault("Page", page_no)
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
bracings: [{mark, section_size, length_mm, length_method}]

Rules:
- Use marks exactly as on the drawing (B1, C3, BP2, BR1).
- Section sizes like W18x35, ISMB400, UB457x191x67.
- Lengths/heights in millimetres (convert metres ×1000). Convert cm ×10.
- PLAN LENGTH RULE: the size/dimension is written in the SAME DIRECTION as the beam.
  Vertical beams use the vertical dimension string beside them (e.g. B3=1500, B8=6000, B7=2000).
  Horizontal beams use the horizontal dimension string (e.g. B4=6000).
  If the same mark appears twice (two B8), each has its own parallel dimension (often both 6000).
- DIAGONAL RULE: if a beam/column/brace is placed diagonally (e.g. BR1, BR4),
  compute length with the triangle diagonal formula L = sqrt(a^2 + b^2)
  where a and b are the horizontal and vertical bay spans the member covers.
  Set length_method to e.g. "sqrt(3000^2+2000^2)".
- If a field is unknown, use empty string.
- Return ONLY valid JSON, no markdown.
"""


def ai_extract_from_text(text: str, page_no: int) -> dict[str, list[dict]]:
    """
    Call OpenAI (via langchain) to enrich extraction.
    Returns empty lists if no API key or on failure — regex results still apply.
    Prints the raw AI response to the console for debugging.
    """
    empty = {"beams": [], "columns": [], "base_plates": [], "bracings": []}
    api_key = _refresh_openai_key()
    if not api_key or not text.strip():
        reason = (
            "OPENAI_API_KEY not set in .env — paste your sk-... key and restart"
            if not api_key
            else "empty page text"
        )
        print(f"[AI extract] SKIPPED page {page_no} — {reason}")
        return empty
    try:
        from langchain_openai import ChatOpenAI
        from langchain_core.messages import HumanMessage, SystemMessage
        import json

        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        llm = ChatOpenAI(
            model=model_name,
            temperature=0,
            api_key=api_key,
        )
        # Cap text to keep token usage reasonable on huge sheets
        clipped = text[:12000]
        print(f"\n{'=' * 72}")
        print(f"[AI extract] CALLING model={model_name} page={page_no} chars={len(clipped)}")
        print(f"{'=' * 72}")
        resp = llm.invoke(
            [
                SystemMessage(content=AI_SYSTEM_PROMPT),
                HumanMessage(
                    content=f"Page {page_no} drawing text:\n\n{clipped}\n\nExtract members as JSON."
                ),
            ]
        )
        content = resp.content if isinstance(resp.content, str) else str(resp.content)

        # --- Console dump: raw AI return ---
        print(f"\n[AI extract] RAW RESPONSE page {page_no} ({len(content)} chars):")
        print("-" * 72)
        print(content)
        print("-" * 72)

        # Strip optional ```json fences
        cleaned = re.sub(r"^```(?:json)?\s*", "", content.strip())
        cleaned = re.sub(r"\s*```$", "", cleaned)

        data = json.loads(cleaned)
        parsed = {
            "beams": data.get("beams") or [],
            "columns": data.get("columns") or [],
            "base_plates": data.get("base_plates") or data.get("basePlates") or [],
            "bracings": data.get("bracings") or [],
        }
        print(
            f"[AI extract] PARSED page {page_no}: "
            f"beams={len(parsed['beams'])} columns={len(parsed['columns'])} "
            f"base_plates={len(parsed['base_plates'])} bracings={len(parsed['bracings'])}"
        )
        print(f"[AI extract] PARSED JSON page {page_no}:")
        print(json.dumps(parsed, indent=2, default=str))
        print(f"{'=' * 72}\n")
        return parsed
    except Exception as exc:  # noqa: BLE001
        print(f"[AI extract] ERROR page {page_no}: {exc}")
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
                "Length Method": r.get("length_method") or "",
                "Page": page_no,
            }
        )
    return out


def dedupe_by_mark(rows: list[dict[str, Any]], allow_multi: bool = False) -> list[dict[str, Any]]:
    """
    Keep the richest row per Mark (most non-empty fields).
    When allow_multi=True (plan beams/bracings), keep separate instances of the
    same mark (e.g. two B8 members each 6000) keyed by Mark+Length+Page.
    """
    if not allow_multi:
        best: dict[str, dict[str, Any]] = {}
        for row in rows:
            mark = row.get("Mark") or ""
            if not mark:
                continue
            score = sum(1 for v in row.values() if v not in ("", None))
            if mark not in best or score > sum(
                1 for v in best[mark].values() if v not in ("", None)
            ):
                best[mark] = row
        return list(best.values())

    best: dict[str, dict[str, Any]] = {}
    for row in rows:
        mark = row.get("Mark") or ""
        if not mark:
            continue
        length = row.get("Length (mm)", row.get("Height (mm)", ""))
        page = row.get("Page", "")
        inst = row.get("_instance", "")
        key = f"{mark}|{length}|{page}|{row.get('Length Method', '')}|{inst}"
        score = sum(1 for k, v in row.items() if k != "_instance" and v not in ("", None))
        if key not in best or score > sum(
            1 for k, v in best[key].items() if k != "_instance" and v not in ("", None)
        ):
            best[key] = row
    # Strip internal instance keys before returning
    cleaned = []
    for row in best.values():
        r = dict(row)
        r.pop("_instance", None)
        cleaned.append(r)
    return cleaned


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

    Length association on plan pages:
      - Orthogonal members → dimension written in the same direction as the member
      - Diagonal members (BR*) → L = √(a² + b²) from bay width × height
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

            # Word-coordinate geometry (digital PDFs) for direction-aware lengths
            geometry = collect_page_geometry(page) if source.startswith("digital") else {
                "beams": [],
                "columns": [],
                "bracings": [],
                "base_plates": [],
                "dims": [],
            }
            # Graphics-layer dims (1500/2000/6000 etc.) via OCR — needed when the
            # PDF text layer only has marks, not dimension strings.
            if len(geometry.get("dims") or []) < 8:
                ocr_dims = ocr_plan_dimensions(page)
                geometry["dims"] = (geometry.get("dims") or []) + ocr_dims
            scale = estimate_scale_mm_per_unit(page)

            # Regex + geometry pass
            if page_no in plan_set:
                page_beams = extract_beams_regex(
                    text, page_no, geometry=geometry, scale=scale
                )
                page_beams = _apply_diagonal_override_for_sloping_beams(
                    page_beams, text, geometry
                )
                beams.extend(page_beams)
                bracings.extend(
                    extract_bracings_regex(
                        text, page_no, geometry=geometry, scale=scale
                    )
                )
            if page_no in elev_set:
                columns.extend(
                    extract_columns_regex(
                        text, page_no, geometry=geometry, scale=scale
                    )
                )
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

    beams = dedupe_by_mark(beams, allow_multi=True)
    columns = dedupe_by_mark(columns, allow_multi=False)
    base_plates = dedupe_by_mark(base_plates, allow_multi=False)
    bracings = dedupe_by_mark(bracings, allow_multi=True)

    summary_rows, engineer_notes = build_summary(
        beams, columns, base_plates, bracings, page_sources, all_text_snippets
    )

    # Structured takeoff lines for UI / Excel consumers
    mark_quantity = {
        "beams": [
            {
                "sentence": _mark_quantity_sentence("Beam", m, c, L, "length"),
                "mark": m,
                "length": L,
                "quantity": c,
            }
            for m, L, c in _count_by_mark_and_length(beams, "Length (mm)")
        ],
        "columns": [
            {
                "sentence": _mark_quantity_sentence("Column", m, c, L, "length"),
                "mark": m,
                "length": L,
                "quantity": c,
            }
            for m, L, c in _count_by_mark_and_length(columns, "Height (mm)")
        ],
        "bracings": [
            {
                "sentence": _mark_quantity_sentence("Bracing", m, c, L, "length"),
                "mark": m,
                "length": L,
                "quantity": c,
            }
            for m, L, c in _count_by_mark_and_length(bracings, "Length (mm)")
        ],
        "base_plates": [
            {
                "sentence": _mark_quantity_sentence("Base plate", m, c, w, "weight"),
                "mark": m,
                "weight": w,
                "quantity": c,
            }
            for m, w, c in _count_by_mark_and_length(base_plates, "Weight (kg)")
        ],
    }

    # Print takeoff sentences to console (same style as Excel Summary)
    print(f"\n{'=' * 72}")
    print("[SUMMARY] Mark × Length × Quantity")
    print(f"{'=' * 72}")
    for group in ("beams", "columns", "bracings", "base_plates"):
        items = mark_quantity.get(group) or []
        if not items:
            continue
        print(f"\n{group.upper()}:")
        for item in items:
            print(f"  {item['sentence']}")
    print(f"{'=' * 72}\n")

    return {
        "beams": beams,
        "columns": columns,
        "base_plates": base_plates,
        "bracings": bracings,
        "summary": summary_rows,
        "engineer_notes": engineer_notes,
        "mark_quantity": mark_quantity,
        "page_sources": page_sources,
    }


def _apply_diagonal_override_for_sloping_beams(
    beams: list[dict[str, Any]],
    text: str,
    geometry: dict[str, list],
) -> list[dict[str, Any]]:
    """
    If text near a beam mark says DIAGONAL / SLOPING / INCLINED, recompute
    length with √(a² + b²) using nearby bay dimensions.
    Uses whole-word mark matching so B4 does not match inside BR4.
    """
    dims = geometry.get("dims") or []
    # Keep ALL geometry instances per mark (two B8s, etc.)
    geo_items = geometry.get("beams") or []
    out = []
    for row in beams:
        mark = row.get("Mark") or ""
        local = _mark_local_text(text, mark, radius=40)
        # Only treat as diagonal when the keyword is on THIS mark's compact window
        is_diag = bool(
            re.search(r"\b(DIAGONAL|SLOPING|INCLINED|SLOPED)\b", local, re.IGNORECASE)
        )
        if is_diag and dims:
            # Prefer the geometry item closest to this row's prior coords if any;
            # otherwise first geometry match for the mark.
            g = next((x for x in geo_items if x.get("mark") == mark), None)
            if g is not None:
                length, how = diagonal_length_from_bay(g, dims, local)
                if length is not None:
                    row = dict(row)
                    row["Length (mm)"] = length
                    row["Length Method"] = how
        out.append(row)
    return out


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


def _format_length_display(length: Any) -> str:
    """Pretty-print length for summary sentences (1500 not 1500.0)."""
    if length in ("", None, "UNKNOWN"):
        return ""
    try:
        f = float(length)
        return str(int(f)) if f == int(f) else str(f)
    except (TypeError, ValueError):
        return str(length)


def _count_by_mark_and_length(
    rows: list[dict],
    length_key: str = "Length (mm)",
) -> list[tuple[str, str, int]]:
    """
    Group members by Mark + Length → quantity.
    Returns list of (mark, length_display, quantity) sorted by mark then length.
    Example: B3@1500×4, B3@2000×8, B3@6000×2, B7@2000×76
    """
    counter: Counter = Counter()
    for r in rows:
        mark = str(r.get("Mark") or "UNKNOWN")
        length = r.get(length_key)
        if length in ("", None):
            length_key_str = "UNKNOWN"
        else:
            length_key_str = _format_length_display(length) or "UNKNOWN"
        counter[(mark, length_key_str)] += 1

    def _sort_key(item: tuple[tuple[str, str], int]) -> tuple:
        (mark, length), _cnt = item
        try:
            length_num = float(length) if length != "UNKNOWN" else -1
        except ValueError:
            length_num = -1
        return (mark, length_num)

    return [
        (mark, length, cnt)
        for (mark, length), cnt in sorted(counter.items(), key=_sort_key)
    ]


def _mark_quantity_sentence(
    kind: str,
    mark: str,
    quantity: int,
    length: str = "",
    unit: str = "length",
) -> str:
    """
    Human summary line matching takeoff style:
      Beam B3 are 4 of length 1500
      Beam B7 are 76 of length 2000
      Base plate BP1 are 11
    """
    if length and length != "UNKNOWN":
        return f"{kind} {mark} are {quantity} of {unit} {length}"
    return f"{kind} {mark} are {quantity}"


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
    - Mark × length × quantity sentences, e.g.:
        Beam B3 are 4 of length 1500
        Beam B3 are 8 of length 2000
        Beam B7 are 76 of length 2000
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

    # --- Mark × Length × Quantity (primary takeoff view) ---
    beam_mq = _count_by_mark_and_length(beams, "Length (mm)")
    col_mq = _count_by_mark_and_length(columns, "Height (mm)")
    brace_mq = _count_by_mark_and_length(bracings, "Length (mm)")
    bp_mq = _count_by_mark_and_length(base_plates, "Weight (kg)")

    for mark, length, cnt in beam_mq:
        sentence = _mark_quantity_sentence("Beam", mark, cnt, length, "length")
        add("Beam Mark Qty", sentence, cnt)
    for mark, length, cnt in col_mq:
        sentence = _mark_quantity_sentence("Column", mark, cnt, length, "length")
        add("Column Mark Qty", sentence, cnt)
    for mark, length, cnt in brace_mq:
        sentence = _mark_quantity_sentence("Bracing", mark, cnt, length, "length")
        add("Bracing Mark Qty", sentence, cnt)
    for mark, weight, cnt in bp_mq:
        sentence = _mark_quantity_sentence("Base plate", mark, cnt, weight, "weight")
        add("Base Plate Mark Qty", sentence, cnt)

    # Size counts
    for size, cnt in _count_by_key(beams, "Section Size"):
        add("Beam Sizes", f"{cnt} beam(s) of size {size}", cnt)
    for size, cnt in _count_by_key(columns, "Section Size"):
        add("Column Sizes", f"{cnt} column(s) of size {size}", cnt)
    for size, cnt in _count_by_key(base_plates, "Plate Size"):
        add("Base Plate Sizes", f"{cnt} base plate(s) of size {size}", cnt)

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

    # Engineer narrative — mark / length / quantity sentences
    notes_lines = [
        "SENIOR STEEL STRUCTURE ENGINEER — DRAWING SUMMARY",
        "=" * 56,
        f"This drawing set contains {len(columns)} column(s), {len(beams)} beam(s), "
        f"{len(bracings)} bracing(s), and {len(base_plates)} base plate(s).",
        "",
    ]

    if beam_mq:
        notes_lines.append("BEAMS (mark × length × quantity):")
        for mark, length, cnt in beam_mq:
            notes_lines.append(
                f"  • {_mark_quantity_sentence('Beam', mark, cnt, length, 'length')}"
            )
        notes_lines.append("")

    if col_mq:
        notes_lines.append("COLUMNS (mark × length × quantity):")
        for mark, length, cnt in col_mq:
            notes_lines.append(
                f"  • {_mark_quantity_sentence('Column', mark, cnt, length, 'length')}"
            )
        notes_lines.append("")

    if brace_mq:
        notes_lines.append("BRACINGS (mark × length × quantity):")
        for mark, length, cnt in brace_mq:
            notes_lines.append(
                f"  • {_mark_quantity_sentence('Bracing', mark, cnt, length, 'length')}"
            )
        notes_lines.append("")

    if bp_mq:
        notes_lines.append("BASE PLATES (mark × weight × quantity):")
        for mark, weight, cnt in bp_mq:
            notes_lines.append(
                f"  • {_mark_quantity_sentence('Base plate', mark, cnt, weight, 'weight')}"
            )
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

    # Optional AI polish — keep mark/length/quantity sentences intact
    engineer_notes = "\n".join(notes_lines)
    api_key = _refresh_openai_key()
    if api_key:
        try:
            from langchain_openai import ChatOpenAI
            from langchain_core.messages import HumanMessage, SystemMessage

            model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
            llm = ChatOpenAI(
                model=model_name,
                temperature=0.2,
                api_key=api_key,
            )
            print(f"\n{'=' * 72}")
            print(f"[AI summary] CALLING model={model_name}")
            print(f"{'=' * 72}")
            polish = llm.invoke(
                [
                    SystemMessage(
                        content=(
                            "You are a senior structural steel engineer. "
                            "Rewrite into a concise professional summary. "
                            "KEEP every mark/length/quantity sentence exactly, e.g. "
                            "'Beam B3 are 4 of length 1500', "
                            "'Beam B7 are 76 of length 2000'. Plain text only."
                        )
                    ),
                    HumanMessage(content=engineer_notes[:8000]),
                ]
            )
            polished = polish.content if isinstance(polish.content, str) else str(polish.content)
            print(f"\n[AI summary] RAW RESPONSE ({len(polished)} chars):")
            print("-" * 72)
            print(polished)
            print("-" * 72)
            print(f"{'=' * 72}\n")
            if polished.strip():
                engineer_notes = polished.strip()
        except Exception as exc:  # noqa: BLE001
            print(f"[AI summary] ERROR: {exc}")
    else:
        print("[AI summary] SKIPPED — OPENAI_API_KEY not set")

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


def _beams_takeoff_rows(beams: list[dict]) -> list[dict[str, Any]]:
    """
    Beams Excel sheet: only Mark, Quantity, Length — one row per mark+length.
    Example:
      B3 | 4 | 1500
      B3 | 8 | 2000
      B3 | 2 | 6000
      B7 | 76 | 2000
    """
    grouped = _count_by_mark_and_length(beams, "Length (mm)")
    return [
        {
            "Mark": mark,
            "Quantity": qty,
            "Length": length if length != "UNKNOWN" else "",
        }
        for mark, length, qty in grouped
    ]


def build_excel(result: dict[str, Any]) -> bytes:
    wb = Workbook()

    # Sheet 1 — Beams (Mark / Quantity / Length only)
    ws1 = wb.active
    ws1.title = "Beams"
    _write_sheet(
        ws1,
        _beams_takeoff_rows(result.get("beams") or []),
        ["Mark", "Quantity", "Length"],
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
    # Append bracing table under narrative for transparency
    brace_start = len(result["summary"]) + 16
    ws4.cell(
        row=brace_start,
        column=1,
        value="Bracing Lengths (√(a²+b²) when diagonal)",
    ).font = Font(bold=True)
    brace_rows = result.get("bracings") or []
    if brace_rows:
        headers = ["Mark", "Section Size", "Length (mm)", "Length Method", "Page"]
        for c_idx, h in enumerate(headers, start=1):
            cell = ws4.cell(row=brace_start + 1, column=c_idx, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        for r_idx, br in enumerate(brace_rows, start=brace_start + 2):
            ws4.cell(row=r_idx, column=1, value=br.get("Mark", ""))
            ws4.cell(row=r_idx, column=2, value=br.get("Section Size", ""))
            ws4.cell(row=r_idx, column=3, value=br.get("Length (mm)", ""))
            ws4.cell(row=r_idx, column=4, value=br.get("Length Method", ""))
            ws4.cell(row=r_idx, column=5, value=br.get("Page", ""))

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
    key = _refresh_openai_key()
    return {
        "status": "ok",
        "app": "SteelDraw AI Extractor",
        "openai_configured": bool(key),
        "openai_model": os.getenv("OPENAI_MODEL", "gpt-4o-mini") if key else None,
        "env_files": {
            "root": str(_ROOT_ENV),
            "root_exists": _ROOT_ENV.exists(),
            "backend": str(_BACKEND_ENV),
            "backend_exists": _BACKEND_ENV.exists(),
        },
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
                "mark_quantity": result.get("mark_quantity") or {},
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
