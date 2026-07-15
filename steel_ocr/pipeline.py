"""End-to-end pipeline: every upload is scanned fresh (OCR + optional vision).

No hardcoded quantities or sizes. Each PDF/image is measured independently:
  1. Digital PDF text layer mark counts (when present)
  2. High-capacity tiled OCR for marks + nearby dimensions
  3. Optional vision LLM (OpenAI / Gemini) to read like a human detailer
  4. Bracing length L = √(a² + b²) from nearby run × rise
  5. Same mark at different lengths → separate Excel rows
"""

from __future__ import annotations

import base64
import collections
import logging
import os
import re
from pathlib import Path
from typing import Any

from pdf2image import convert_from_path
from PIL import Image

from steel_ocr.excel_export import build_excel_bytes, members_preview
from steel_ocr.extractors import extract_document_members
from steel_ocr.reader import DocumentText, read_document
from steel_ocr.spatial import (
    BEAM_RE,
    BP_RE,
    BRACE_RE,
    OCR_DPI,
    PB_RE,
    PLAT_RE,
    apply_spatial_sizes,
    infer_size_quantities_from_plan_image,
    ocr_tokens_with_boxes,
    scale_size_counts_to_total,
)
from steel_ocr.vision import (
    classify_vision_buckets,
    merge_ocr_and_vision,
    vision_available,
    vision_extract_members,
)

logger = logging.getLogger("steel_ocr.pipeline")

BEAM_TOKEN_RE = re.compile(r"^B[1-9]\d*$")
BRACE_TOKEN_RE = re.compile(r"^BR\d+$")
BP_TOKEN_RE = re.compile(r"^BP\d+$")
PB_TOKEN_RE = re.compile(r"^PB\d+$")
PLAT_TOKEN_RE = re.compile(r"^PB[A-G]$|^PAB$")


def _count_plan_marks_from_pdf(path: Path) -> dict[str, collections.Counter]:
    """Count mark tokens on this PDF's digital text layer (this file only)."""
    import pdfplumber

    buckets = {
        "beams": collections.Counter(),
        "braces": collections.Counter(),
        "base_plates": collections.Counter(),
        "pb": collections.Counter(),
        "platforms": collections.Counter(),
    }
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for w in page.extract_words() or []:
                t = w["text"]
                if BEAM_TOKEN_RE.fullmatch(t):
                    buckets["beams"][t] += 1
                elif BRACE_TOKEN_RE.fullmatch(t):
                    buckets["braces"][t] += 1
                elif BP_TOKEN_RE.fullmatch(t):
                    buckets["base_plates"][t] += 1
                elif PB_TOKEN_RE.fullmatch(t):
                    buckets["pb"][t] += 1
                elif PLAT_TOKEN_RE.fullmatch(t):
                    buckets["platforms"][t] += 1
    return buckets


def _count_marks_from_ocr_tokens(
    marks: list[dict[str, Any]],
) -> dict[str, collections.Counter]:
    buckets = {
        "beams": collections.Counter(),
        "braces": collections.Counter(),
        "base_plates": collections.Counter(),
        "pb": collections.Counter(),
        "platforms": collections.Counter(),
    }
    for m in marks:
        t = m["text"]
        if BEAM_RE.fullmatch(t):
            buckets["beams"][t] += 1
        elif BRACE_RE.fullmatch(t):
            buckets["braces"][t] += 1
        elif BP_RE.fullmatch(t):
            buckets["base_plates"][t] += 1
        elif PB_RE.fullmatch(t):
            buckets["pb"][t] += 1
        elif PLAT_RE.fullmatch(t):
            buckets["platforms"][t] += 1
    return buckets


def _merge_mark_counts(
    digital: collections.Counter,
    ocr: collections.Counter,
) -> collections.Counter:
    """
    Prefer larger of digital vs OCR per mark; keep OCR-only marks (e.g. BR4).

    Drops OCR-only garbage like B38 / B67 when they look like concatenated
    neighbours of marks already present on the digital layer.
    """
    merged: collections.Counter = collections.Counter()
    dig_names = set(digital)
    for name in set(digital) | set(ocr):
        d = int(digital.get(name, 0))
        o = int(ocr.get(name, 0))
        if d == 0 and o > 0 and dig_names and _looks_like_ocr_concat(name, dig_names):
            logger.info("Dropping OCR noise mark %s (not in text layer)", name)
            continue
        merged[name] = max(d, o)
    return merged


def _looks_like_ocr_concat(name: str, known: set[str]) -> bool:
    """True if name looks like two known marks glued (B3+B8 → B38)."""
    if not re.fullmatch(r"B\d{2,3}", name):
        return False
    # B38 → B3 + 8 or B + 38
    body = name[1:]
    for i in range(1, len(body)):
        left, right = "B" + body[:i], "B" + body[i:]
        if left in known and (right in known or body[i:].isdigit()):
            if left in known:
                return True
    return False


def _render_plan_images(path: Path, dpi: int) -> list[Image.Image]:
    """Render this upload for OCR/vision — almost full sheet (light title crop)."""
    suffix = path.suffix.lower()
    images: list[Image.Image] = []
    crop_right = float(os.getenv("OCR_CROP_RIGHT", "0.92"))  # keep more of sheet
    crop_bottom = float(os.getenv("OCR_CROP_BOTTOM", "0.95"))
    if suffix == ".pdf":
        for im in convert_from_path(str(path), dpi=dpi):
            w, h = im.size
            images.append(im.crop((0, 0, int(w * crop_right), int(h * crop_bottom))))
    else:
        images.append(Image.open(path))
    return images


def _ocr_plan_takeoff(
    images: list[Image.Image],
) -> tuple[dict[str, collections.Counter], dict[str, collections.Counter]]:
    mark_counts = {
        "beams": collections.Counter(),
        "braces": collections.Counter(),
        "base_plates": collections.Counter(),
        "pb": collections.Counter(),
        "platforms": collections.Counter(),
    }
    size_qty: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)

    for im in images:
        marks, _dims = ocr_tokens_with_boxes(im, high_capacity=True)
        page_marks = _count_marks_from_ocr_tokens(marks)
        for key in mark_counts:
            mark_counts[key].update(page_marks[key])
        for name, ctr in infer_size_quantities_from_plan_image(im, high_capacity=True).items():
            size_qty[name].update(ctr)

    return mark_counts, dict(size_qty)


def _rows_from_mark_and_sizes(
    mark_totals: collections.Counter,
    size_qty: dict[str, collections.Counter],
    *,
    page: int = 1,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for name in sorted(mark_totals, key=lambda x: (len(x), x)):
        total = int(mark_totals[name])
        if total <= 0:
            continue
        sizes = size_qty.get(name) or collections.Counter()
        if not sizes:
            rows.append(
                {
                    "Member Name": name,
                    "Quantity": total,
                    "Size": "",
                    "Length": "",
                    "Length Note": "",
                    "Page": page,
                    "Source": "scan",
                }
            )
            continue
        allocated = scale_size_counts_to_total(sizes, total)
        for raw_size, qty in sorted(
            allocated.items(),
            key=lambda kv: (
                int(re.match(r"^\d+", kv[0]).group(0)) if re.match(r"^\d+", kv[0]) else 0,
                kv[0],
            ),
        ):
            if qty <= 0:
                continue
            size, note = raw_size, ""
            if "|" in raw_size:
                size, note = raw_size.split("|", 1)
            rows.append(
                {
                    "Member Name": name,
                    "Quantity": int(qty),
                    "Size": size,
                    "Length": size,
                    "Length Note": note,
                    "Page": page,
                    "Source": "scan",
                }
            )
    return rows


def process_drawing(
    path: str | Path,
    *,
    dpi: int | None = None,
    include_details: bool = True,
    use_vision: bool | None = None,
) -> dict[str, Any]:
    """
    Scan this upload only — no memory of previous drawings, no hardcodes.

    Flow for every PDF/image:
      digital mark counts → high-capacity OCR → optional vision → Excel
    """
    path = Path(path)
    logger.info("Scanning upload: %s (fresh takeoff, no hardcodes)", path.name)
    doc: DocumentText = read_document(path, dpi=dpi)

    digital_counts = (
        _count_plan_marks_from_pdf(path) if path.suffix.lower() == ".pdf" else None
    )
    ocr_dpi = dpi or OCR_DPI
    size_qty: dict[str, collections.Counter] = {}
    ocr_mark_counts = {
        k: collections.Counter()
        for k in ("beams", "braces", "base_plates", "pb", "platforms")
    }
    images: list[Image.Image] = []

    try:
        images = _render_plan_images(path, dpi=ocr_dpi)
        ocr_mark_counts, size_qty = _ocr_plan_takeoff(images)
        logger.info(
            "OCR scan — beams=%s braces=%s | sizes=%s",
            dict(ocr_mark_counts["beams"]),
            dict(ocr_mark_counts["braces"]),
            {k: dict(v) for k, v in size_qty.items()},
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning("OCR scan failed: %s", exc)

    dig = digital_counts or {
        k: collections.Counter()
        for k in ("beams", "braces", "base_plates", "pb", "platforms")
    }
    beam_totals = _merge_mark_counts(dig["beams"], ocr_mark_counts["beams"])
    beam_totals.update(_merge_mark_counts(dig["pb"], ocr_mark_counts["pb"]))
    brace_totals = _merge_mark_counts(dig["braces"], ocr_mark_counts["braces"])
    for name, qty in brace_totals.items():
        beam_totals[name] = max(beam_totals.get(name, 0), qty)
    plate_totals = _merge_mark_counts(dig["base_plates"], ocr_mark_counts["base_plates"])
    col_totals = _merge_mark_counts(dig["platforms"], ocr_mark_counts["platforms"])

    beams = _rows_from_mark_and_sizes(beam_totals, size_qty)
    columns = _rows_from_mark_and_sizes(col_totals, {})
    plates = _rows_from_mark_and_sizes(plate_totals, size_qty)

    # Fallback: schedule regex if scan found almost nothing
    if not beams and not columns and not plates:
        logger.info("Sparse scan — falling back to text regex extractors")
        members = extract_document_members(doc)
        flat = {n: c.most_common(1)[0][0] for n, c in size_qty.items() if c}
        beams = apply_spatial_sizes(members["beams"], flat)
        columns = apply_spatial_sizes(members["columns"], flat)
        plates = apply_spatial_sizes(members["base_plates"], flat)

    # Optional vision — sees this upload like a human; never uses prior drawings
    vision_on = vision_available() if use_vision is None else use_vision
    vision_rows: list[dict[str, Any]] = []
    if vision_on and images:
        vision_rows = vision_extract_members(images)
        if vision_rows:
            combined = merge_ocr_and_vision(beams + columns + plates, vision_rows)
            beams, columns, plates = classify_vision_buckets(combined)

    excel_bytes = build_excel_bytes(
        beams, columns, plates, include_details=include_details
    )
    preview = members_preview(beams, columns, plates)

    return {
        "filename": doc.filename,
        "document": doc.as_dict(),
        "preview": preview,
        "raw": {"beams": beams, "columns": columns, "base_plates": plates},
        "scan": {
            "mode": "ocr+vision" if vision_rows else "ocr",
            "vision_enabled": bool(vision_on and vision_available()),
            "dpi": ocr_dpi,
            "digital": {k: dict(v) for k, v in dig.items()},
            "ocr": {k: dict(v) for k, v in ocr_mark_counts.items()},
            "size_quantities": {k: dict(v) for k, v in size_qty.items()},
        },
        "excel_bytes": excel_bytes,
        "excel_base64": base64.b64encode(excel_bytes).decode("ascii"),
        "excel_filename": f"{path.stem}_steel_takeoff.xlsx",
    }


def process_drawing_to_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    **kwargs: Any,
) -> Path:
    result = process_drawing(input_path, **kwargs)
    out = Path(output_path) if output_path else Path(result["excel_filename"])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(result["excel_bytes"])
    logger.info("Wrote Excel: %s", out)
    return out
