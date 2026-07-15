"""Read text from PDFs, scanned PDFs, and images via digital extract + OCR."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

from steel_ocr.preprocess import pil_to_bgr, preprocess_for_ocr, preprocess_variants

logger = logging.getLogger("steel_ocr.reader")

DIGITAL_TEXT_THRESHOLD = 40
DEFAULT_DPI = int(os.getenv("OCR_DPI", "250"))
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
PDF_EXTENSIONS = {".pdf"}

# Allow override of tesseract binary path
_tesseract_cmd = os.getenv("TESSERACT_CMD", "").strip()
if _tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd


@dataclass
class PageText:
    """Text extracted from one page / image."""

    page_number: int
    text: str
    source: str  # digital | ocr | hybrid
    page_type: str = "Other"  # Plan | Elevation | Other
    width: float | None = None
    height: float | None = None


@dataclass
class DocumentText:
    """All pages from an uploaded drawing."""

    filename: str
    pages: list[PageText] = field(default_factory=list)

    @property
    def full_text(self) -> str:
        parts = []
        for page in self.pages:
            parts.append(f"\n===== PAGE {page.page_number} ({page.source}) =====\n")
            parts.append(page.text or "")
        return "\n".join(parts)

    def as_dict(self) -> dict[str, Any]:
        return {
            "filename": self.filename,
            "page_count": len(self.pages),
            "pages": [
                {
                    "page_number": p.page_number,
                    "source": p.source,
                    "page_type": p.page_type,
                    "char_count": len(p.text or ""),
                    "preview": (p.text or "")[:400],
                }
                for p in self.pages
            ],
        }


def _ocr_pil(image: Image.Image, multi_pass: bool = True) -> str:
    """Run Tesseract on a PIL image with optional multi-pass preprocessing."""
    bgr = pil_to_bgr(image)
    config = "--oem 3 --psm 6"
    texts: list[str] = []

    if multi_pass:
        for variant in preprocess_variants(bgr):
            try:
                texts.append(pytesseract.image_to_string(variant, config=config) or "")
            except Exception as exc:  # noqa: BLE001
                logger.warning("OCR variant failed: %s", exc)
    else:
        processed = preprocess_for_ocr(bgr)
        texts.append(pytesseract.image_to_string(processed, config=config) or "")

    # Pick the richest OCR pass (most alphanumeric characters)
    def score(t: str) -> int:
        return sum(ch.isalnum() for ch in t)

    best = max(texts, key=score) if texts else ""
    return best.strip()


def _page_has_digital_text(page: pdfplumber.page.Page) -> bool:
    try:
        text = page.extract_text() or ""
        return len(text.strip()) >= DIGITAL_TEXT_THRESHOLD
    except Exception:  # noqa: BLE001
        return False


def detect_page_type(text: str) -> str:
    """Classify sheet as Plan, Elevation, or Other from title-block keywords."""
    import re

    head = (text or "")[:2500]
    elev_re = re.compile(
        r"\b(?:ELEVATION|BRACE\s+FRAME\s+ELEV|FRAME\s+ELEVATION|"
        r"BUILDING\s+ELEVATION|SECTION\s+ELEVATION)\b",
        re.IGNORECASE,
    )
    plan_re = re.compile(
        r"\b(?:FLOOR\s+FRAMING|FRAMING\s+PLAN|ROOF\s+PLAN|FOUNDATION\s+PLAN|"
        r"STRUCTURAL\s+PLAN|PLAN\s*[-–]|PLAN\b)\b",
        re.IGNORECASE,
    )
    elev_hit = elev_re.search(head)
    plan_hit = plan_re.search(head)
    if elev_hit and not plan_hit:
        return "Elevation"
    if plan_hit and not elev_hit:
        return "Plan"
    if elev_hit and plan_hit:
        return "Elevation" if elev_hit.start() < plan_hit.start() else "Plan"
    return "Other"


def read_image_file(path: str | Path, page_number: int = 1) -> PageText:
    """OCR a standalone image file (PNG/JPG/TIFF/…)."""
    path = Path(path)
    image = Image.open(path)
    text = _ocr_pil(image, multi_pass=True)
    return PageText(
        page_number=page_number,
        text=text,
        source="ocr",
        page_type=detect_page_type(text),
        width=float(image.width),
        height=float(image.height),
    )


def read_pdf_file(path: str | Path, dpi: int | None = None) -> list[PageText]:
    """
    Extract text from every PDF page.

    Digital pages use pdfplumber text. Scanned / image-only pages (and pages
    with very little text) are rendered and OCR'd. Hybrid: if digital text is
    thin, OCR is also run and the richer result is kept.
    """
    path = Path(path)
    dpi = dpi or DEFAULT_DPI
    pages: list[PageText] = []

    with pdfplumber.open(str(path)) as pdf:
        total = len(pdf.pages)
        # Pre-render OCR candidates only when needed (lazy per page)
        rendered: dict[int, Image.Image] = {}

        def get_render(idx0: int) -> Image.Image:
            if idx0 not in rendered:
                images = convert_from_path(
                    str(path),
                    dpi=dpi,
                    first_page=idx0 + 1,
                    last_page=idx0 + 1,
                )
                rendered[idx0] = images[0] if images else Image.new("RGB", (100, 100), "white")
            return rendered[idx0]

        for idx, page in enumerate(pdf.pages):
            page_num = idx + 1
            digital = ""
            try:
                digital = (page.extract_text() or "").strip()
            except Exception as exc:  # noqa: BLE001
                logger.warning("Digital extract failed page %s: %s", page_num, exc)

            source = "digital"
            text = digital

            needs_ocr = len(digital) < DIGITAL_TEXT_THRESHOLD
            # Also OCR when page looks image-heavy even if some text exists
            if not needs_ocr and page.images and len(digital) < 200:
                needs_ocr = True

            if needs_ocr:
                try:
                    ocr_text = _ocr_pil(get_render(idx), multi_pass=True)
                    if len(ocr_text) > len(digital):
                        text = ocr_text
                        source = "ocr" if not digital else "hybrid"
                    elif digital and ocr_text:
                        # Merge unique lines from OCR into digital
                        merged = digital + "\n" + ocr_text
                        text = merged
                        source = "hybrid"
                    elif ocr_text:
                        text = ocr_text
                        source = "ocr"
                except Exception as exc:  # noqa: BLE001
                    logger.warning("OCR failed page %s/%s: %s", page_num, total, exc)

            pages.append(
                PageText(
                    page_number=page_num,
                    text=text or "",
                    source=source,
                    page_type=detect_page_type(text or ""),
                    width=float(page.width) if page.width else None,
                    height=float(page.height) if page.height else None,
                )
            )

    return pages


def read_document(path: str | Path, dpi: int | None = None) -> DocumentText:
    """
    Load a steel drawing from PDF or image and return page texts.

    Supported:
      - .pdf (digital, scanned, or mixed image+text)
      - .png .jpg .jpeg .tif .tiff .bmp .webp
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Drawing not found: {path}")

    suffix = path.suffix.lower()
    if suffix in PDF_EXTENSIONS:
        pages = read_pdf_file(path, dpi=dpi)
    elif suffix in IMAGE_EXTENSIONS:
        pages = [read_image_file(path, page_number=1)]
    else:
        raise ValueError(
            f"Unsupported file type '{suffix}'. "
            f"Use PDF or image ({', '.join(sorted(IMAGE_EXTENSIONS))})."
        )

    return DocumentText(filename=path.name, pages=pages)
