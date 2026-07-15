"""End-to-end pipeline: load drawing → OCR → extract → Excel."""

from __future__ import annotations

import base64
import collections
import logging
import re
from pathlib import Path
from typing import Any

from pdf2image import convert_from_path
from PIL import Image

from steel_ocr.excel_export import build_excel_bytes, members_preview
from steel_ocr.extractors import extract_document_members
from steel_ocr.reader import DocumentText, read_document
from steel_ocr.spatial import apply_spatial_sizes, infer_sizes_from_plan_image

logger = logging.getLogger("steel_ocr.pipeline")

BEAM_TOKEN_RE = re.compile(r"^B[1-9]\d*$")
BRACE_TOKEN_RE = re.compile(r"^BR\d+$")
BP_TOKEN_RE = re.compile(r"^BP\d+$")
PB_TOKEN_RE = re.compile(r"^PB\d+$")
PLAT_TOKEN_RE = re.compile(r"^PB[A-G]$|^PAB$")


def _count_plan_marks_from_pdf(path: Path) -> dict[str, collections.Counter]:
    """Count exact mark tokens on digital PDF text layer (reliable for GA plans)."""
    import pdfplumber

    beams: collections.Counter = collections.Counter()
    braces: collections.Counter = collections.Counter()
    plates: collections.Counter = collections.Counter()
    pbs: collections.Counter = collections.Counter()
    plats: collections.Counter = collections.Counter()

    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for w in page.extract_words() or []:
                t = w["text"]
                if BEAM_TOKEN_RE.fullmatch(t):
                    beams[t] += 1
                elif BRACE_TOKEN_RE.fullmatch(t):
                    braces[t] += 1
                elif BP_TOKEN_RE.fullmatch(t):
                    plates[t] += 1
                elif PB_TOKEN_RE.fullmatch(t):
                    pbs[t] += 1
                elif PLAT_TOKEN_RE.fullmatch(t):
                    plats[t] += 1

    return {
        "beams": beams,
        "braces": braces,
        "base_plates": plates,
        "pb": pbs,
        "platforms": plats,
    }


def _members_from_counts(
    counts: collections.Counter,
    *,
    sizes: dict[str, str] | None = None,
    page: int = 1,
) -> list[dict[str, Any]]:
    sizes = sizes or {}
    rows: list[dict[str, Any]] = []
    for name, qty in sorted(counts.items(), key=lambda x: (len(x[0]), x[0])):
        raw = sizes.get(name, "")
        size, note = raw, ""
        if "|" in raw:
            size, note = raw.split("|", 1)
        rows.append(
            {
                "Member Name": name,
                "Quantity": int(qty),
                "Size": size,
                "Length": size,
                "Length Note": note,
                "Page": page,
            }
        )
    return rows


def _spatial_sizes_for_path(path: Path, dpi: int = 200) -> dict[str, str]:
    """Render plan pages and OCR dimensions written near member marks."""
    suffix = path.suffix.lower()
    images: list[Image.Image] = []
    try:
        if suffix == ".pdf":
            pages = convert_from_path(str(path), dpi=dpi)
            for im in pages:
                # Drop right title-block strip on wide GA sheets
                w, h = im.size
                images.append(im.crop((0, 0, int(w * 0.72), int(h * 0.88))))
        else:
            images.append(Image.open(path))
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not render for spatial OCR: %s", exc)
        return {}

    merged: dict[str, str] = {}
    for im in images:
        try:
            inferred = infer_sizes_from_plan_image(im)
            merged.update(inferred)
        except Exception as exc:  # noqa: BLE001
            logger.warning("spatial size inference failed: %s", exc)
    return merged


def _looks_like_ga_plan(doc: DocumentText, plan_counts: dict[str, collections.Counter]) -> bool:
    """GA plans have many repeated B# marks but little schedule section text."""
    beam_marks = sum(plan_counts["beams"].values())
    text = doc.full_text.upper()
    has_schedule = "SCHEDULE" in text and ("SECTION" in text or "ISMB" in text or "W12" in text)
    return beam_marks >= 8 and not has_schedule


def process_drawing(
    path: str | Path,
    *,
    dpi: int | None = None,
    include_details: bool = True,
) -> dict[str, Any]:
    """
    Full takeoff pipeline for one PDF or image.

    For GA / framing plans, counts mark occurrences on the sheet and uses OCR to
    read dimension text written near each member (e.g. B4 next to 6000).
    """
    path = Path(path)
    logger.info("Reading drawing: %s", path.name)
    doc: DocumentText = read_document(path, dpi=dpi)

    logger.info(
        "Extracted %s page(s); sources=%s",
        len(doc.pages),
        [p.source for p in doc.pages],
    )

    members = extract_document_members(doc)
    beams = members["beams"]
    columns = members["columns"]
    plates = members["base_plates"]

    spatial_sizes: dict[str, str] = {}
    plan_counts: dict[str, collections.Counter] | None = None

    if path.suffix.lower() == ".pdf":
        plan_counts = _count_plan_marks_from_pdf(path)

    # Prefer plan mark counts + nearby OCR sizes for GA drawings
    if plan_counts and _looks_like_ga_plan(doc, plan_counts):
        logger.info("GA plan detected — counting marks + OCR nearby sizes")
        spatial_sizes = _spatial_sizes_for_path(path, dpi=dpi or 200)
        logger.info("Spatial sizes: %s", spatial_sizes)

        beams = _members_from_counts(plan_counts["beams"], sizes=spatial_sizes)
        # Secondary PB / bracing treated as beam-sheet members
        beams.extend(_members_from_counts(plan_counts["pb"], sizes=spatial_sizes))
        beams.extend(_members_from_counts(plan_counts["braces"], sizes=spatial_sizes))

        columns = _members_from_counts(plan_counts["platforms"], sizes={})
        plates = _members_from_counts(plan_counts["base_plates"], sizes=spatial_sizes)
    else:
        # Schedule-style drawings: fill any missing sizes from spatial OCR
        spatial_sizes = _spatial_sizes_for_path(path, dpi=dpi or 200)
        beams = apply_spatial_sizes(beams, spatial_sizes)
        columns = apply_spatial_sizes(columns, spatial_sizes)
        plates = apply_spatial_sizes(plates, spatial_sizes)

    excel_bytes = build_excel_bytes(
        beams, columns, plates, include_details=include_details
    )
    preview = members_preview(beams, columns, plates)

    return {
        "filename": doc.filename,
        "document": doc.as_dict(),
        "preview": preview,
        "raw": {"beams": beams, "columns": columns, "base_plates": plates},
        "spatial_sizes": spatial_sizes,
        "excel_bytes": excel_bytes,
        "excel_base64": base64.b64encode(excel_bytes).decode("ascii"),
        "excel_filename": f"{path.stem}_steel_takeoff.xlsx",
    }


def process_drawing_to_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    **kwargs: Any,
) -> Path:
    """Process a drawing and write the Excel workbook to disk."""
    result = process_drawing(input_path, **kwargs)
    out = Path(output_path) if output_path else Path(result["excel_filename"])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(result["excel_bytes"])
    logger.info("Wrote Excel: %s", out)
    return out
