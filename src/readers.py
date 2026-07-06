"""Readers that turn drawing files into raw text and tables.

- PDF: text and tables via ``pdfplumber``; scanned pages fall back to OpenCV
  pre-processing plus Tesseract OCR.
- DWG/DXF: text entities via ``ezdxf`` (DWG requires the ODA File Converter).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)

SUPPORTED_SUFFIXES = {".pdf", ".dwg", ".dxf"}
# Below this many characters a PDF page is treated as scanned and sent to OCR.
MIN_TEXT_LENGTH = 40


@dataclass
class RawContent:
    """Unified extraction result for a single source file."""

    source: str
    text: str = ""
    tables: list[list[list[str | None]]] = field(default_factory=list)


def read_file(path: Path, use_ocr: bool = True) -> RawContent:
    """Dispatch to the correct reader based on file suffix."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return _read_pdf(path, use_ocr=use_ocr)
    if suffix in (".dwg", ".dxf"):
        return _read_cad(path)
    raise ValueError(f"Unsupported file type: {path.name}")


def _read_pdf(path: Path, use_ocr: bool) -> RawContent:
    import pdfplumber

    texts: list[str] = []
    tables: list[list[list[str | None]]] = []

    with pdfplumber.open(path) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            table_bboxes: list[tuple[float, float, float, float]] = []
            try:
                for found in page.find_tables():
                    extracted = found.extract()
                    if extracted:
                        tables.append(extracted)
                        table_bboxes.append(found.bbox)
            except Exception:  # noqa: BLE001 - malformed tables must not abort
                logger.warning("Failed to extract tables on page %s of %s", index, path.name)

            # Free text = text outside table regions, so table rows are not
            # parsed twice (once as a table and once as flattened text).
            page_text = _text_outside_tables(page, table_bboxes)

            if use_ocr and len(page_text.strip()) < MIN_TEXT_LENGTH and not tables:
                ocr_text = _ocr_page(page, index, path.name)
                if ocr_text:
                    page_text = f"{page_text}\n{ocr_text}".strip()

            texts.append(page_text)

    return RawContent(source=path.name, text="\n".join(texts), tables=tables)


def _text_outside_tables(page, table_bboxes: list[tuple[float, float, float, float]]) -> str:
    if not table_bboxes:
        return page.extract_text() or ""

    def keep(obj) -> bool:
        cx = (obj["x0"] + obj["x1"]) / 2.0
        cy = (obj["top"] + obj["bottom"]) / 2.0
        for x0, top, x1, bottom in table_bboxes:
            if x0 <= cx <= x1 and top <= cy <= bottom:
                return False
        return True

    try:
        return page.filter(keep).extract_text() or ""
    except Exception:  # noqa: BLE001
        return page.extract_text() or ""


def _ocr_page(page, page_number: int, source_name: str) -> str:
    """Run OCR on a rendered PDF page image; returns "" on any failure."""
    try:
        import cv2
        import numpy as np
        import pytesseract
        from PIL import Image
    except Exception as exc:  # noqa: BLE001
        logger.warning("OCR dependencies unavailable (%s); skipping OCR", exc)
        return ""

    try:
        rendered = page.to_image(resolution=200)
        image = np.array(rendered.original.convert("RGB"))
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not render page %s of %s for OCR: %s", page_number, source_name, exc)
        return ""

    try:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        denoised = cv2.fastNlMeansDenoising(gray, h=10)
        _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        logger.info("Running OCR on page %s of %s", page_number, source_name)
        return pytesseract.image_to_string(Image.fromarray(binary), config="--psm 6")
    except Exception as exc:  # noqa: BLE001
        logger.warning("OCR failed on page %s of %s: %s", page_number, source_name, exc)
        return ""


def _read_cad(path: Path) -> RawContent:
    """Extract text entities from a DXF/DWG file using ezdxf."""
    try:
        import ezdxf
    except Exception as exc:  # noqa: BLE001
        logger.warning("ezdxf unavailable (%s); cannot read %s", exc, path.name)
        return RawContent(source=path.name)

    doc = None
    if path.suffix.lower() == ".dxf":
        try:
            doc = ezdxf.readfile(str(path))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to read DXF %s: %s", path.name, exc)
    else:  # .dwg needs the ODA File Converter via the odafc add-on
        try:
            from ezdxf.addons import odafc

            doc = odafc.readfile(str(path))
        except Exception as exc:  # noqa: BLE001
            logger.warning(
                "Could not read DWG %s (ODA File Converter required): %s",
                path.name,
                exc,
            )

    if doc is None:
        return RawContent(source=path.name)

    lines: list[str] = []
    try:
        for entity in doc.modelspace():
            dxftype = entity.dxftype()
            if dxftype == "TEXT":
                lines.append(str(entity.dxf.text))
            elif dxftype == "MTEXT":
                lines.append(entity.plain_text())
    except Exception as exc:  # noqa: BLE001
        logger.warning("Failed while iterating entities in %s: %s", path.name, exc)

    return RawContent(source=path.name, text="\n".join(lines))
