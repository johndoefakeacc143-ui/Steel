"""Unified reader for steel drawings in any format.

Dispatches by file type so the same BOM pipeline works for:
  - text PDFs            -> pdfplumber text/tables + vector-char mark recovery
  - scanned / image PDFs -> page images (OCR applied downstream)
  - standalone images    -> PNG / JPG / JPEG / TIF / TIFF / BMP / WEBP
  - CAD                  -> DXF (and DWG via the ODA File Converter)

Every source is normalised to a list of ``PageContent`` so the rest of the
application does not care what the original format was.
"""

from __future__ import annotations

import logging
from pathlib import Path

import numpy as np

from src.pdf_reader import PageContent, PDFReader

logger = logging.getLogger(__name__)

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
CAD_SUFFIXES = {".dxf", ".dwg"}
SUPPORTED_SUFFIXES = {".pdf"} | IMAGE_SUFFIXES | CAD_SUFFIXES


def read_drawing(path: Path) -> list[PageContent]:
    """Read any supported drawing file into a list of pages."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return PDFReader(path).read()
    if suffix in IMAGE_SUFFIXES:
        return _read_image(path)
    if suffix in CAD_SUFFIXES:
        return _read_cad(path)
    raise ValueError(f"Unsupported file type: {path.name}")


def _read_image(path: Path) -> list[PageContent]:
    """Load an image file as a single page; text is filled in later by OCR."""
    image = _load_image_rgb(path)
    if image is None:
        logger.warning("Could not read image %s", path.name)
        return []
    return [PageContent(page_number=1, text="", tables=[], image=image, stacked_marks={})]


def _load_image_rgb(path: Path) -> np.ndarray | None:
    try:
        from PIL import Image

        with Image.open(path) as img:
            return np.array(img.convert("RGB"))
    except Exception as exc:  # noqa: BLE001
        logger.warning("Failed to load image %s: %s", path.name, exc)
        return None


def _read_cad(path: Path) -> list[PageContent]:
    """Extract text entities from a DXF/DWG (no image, so no dimension OCR)."""
    try:
        import ezdxf
    except Exception as exc:  # noqa: BLE001
        logger.warning("ezdxf unavailable (%s); cannot read %s", exc, path.name)
        return []

    doc = None
    if path.suffix.lower() == ".dxf":
        try:
            doc = ezdxf.readfile(str(path))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to read DXF %s: %s", path.name, exc)
    else:
        try:
            from ezdxf.addons import odafc

            doc = odafc.readfile(str(path))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not read DWG %s (ODA converter required): %s", path.name, exc)

    if doc is None:
        return []

    lines: list[str] = []
    try:
        for entity in doc.modelspace():
            if entity.dxftype() == "TEXT":
                lines.append(str(entity.dxf.text))
            elif entity.dxftype() == "MTEXT":
                lines.append(entity.plain_text())
    except Exception as exc:  # noqa: BLE001
        logger.warning("Failed iterating entities in %s: %s", path.name, exc)

    return [PageContent(page_number=1, text="\n".join(lines), tables=[], image=None, stacked_marks={})]
