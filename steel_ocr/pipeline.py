"""End-to-end pipeline: load drawing → OCR → extract → Excel."""

from __future__ import annotations

import base64
import logging
from pathlib import Path
from typing import Any

from steel_ocr.excel_export import build_excel_bytes, members_preview
from steel_ocr.extractors import extract_document_members
from steel_ocr.reader import DocumentText, read_document

logger = logging.getLogger("steel_ocr.pipeline")


def process_drawing(
    path: str | Path,
    *,
    dpi: int | None = None,
    include_details: bool = True,
) -> dict[str, Any]:
    """
    Full takeoff pipeline for one PDF or image.

    Returns preview tables, raw member lists, document metadata, and Excel bytes.
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

    excel_bytes = build_excel_bytes(
        beams, columns, plates, include_details=include_details
    )
    preview = members_preview(beams, columns, plates)

    return {
        "filename": doc.filename,
        "document": doc.as_dict(),
        "preview": preview,
        "raw": members,
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
