"""Extract text and images from steel structure PDF drawings."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pdfplumber

logger = logging.getLogger(__name__)

# Above this many vector edges on a page, skip pdfplumber table detection
# (it becomes pathologically slow on dense CAD drawings).
MAX_EDGES_FOR_TABLES = 6000


@dataclass
class PageContent:
    page_number: int
    text: str
    tables: list[list[list[str | None]]] = field(default_factory=list)
    image: np.ndarray | None = None
    # Marks recovered from rotated/stacked characters that the text layer misses.
    stacked_marks: dict[str, int] = field(default_factory=dict)


class PDFReader:
    """Read PDF pages using pdfplumber."""

    def __init__(self, pdf_path: Path) -> None:
        self.pdf_path = pdf_path

    def read(self) -> list[PageContent]:
        from src.mark_recovery import recover_stacked_marks

        pages: list[PageContent] = []
        with pdfplumber.open(self.pdf_path) as pdf:
            total = len(pdf.pages)
            for index, page in enumerate(pdf.pages, start=1):
                logger.info("  page %s/%s: extracting text...", index, total)
                text = page.extract_text() or ""

                # pdfplumber's table detection intersects every edge, which is
                # pathologically slow on dense CAD drawings. Skip it when the page
                # has a huge number of edges (those "tables" are the drawing grid,
                # not a BOM schedule, anyway).
                n_edges = len(page.edges)
                if n_edges > MAX_EDGES_FOR_TABLES:
                    logger.warning(
                        "  page %s/%s: %s edges - skipping table detection to avoid a "
                        "very slow scan (use a text-table drawing for schedule parsing)",
                        index,
                        total,
                        n_edges,
                    )
                    tables = []
                else:
                    logger.info("  page %s/%s: extracting tables (%s edges)...", index, total, n_edges)
                    try:
                        tables = page.extract_tables() or []
                    except Exception as exc:  # noqa: BLE001
                        logger.warning("  page %s: table extraction failed: %s", index, exc)
                        tables = []

                logger.info("  page %s/%s: rendering image...", index, total)
                image = self._page_to_image(page)

                logger.info(
                    "  page %s/%s: recovering marks (%s chars)...",
                    index,
                    total,
                    len(page.chars),
                )
                try:
                    stacked = recover_stacked_marks(page.chars, text)
                except Exception as exc:  # noqa: BLE001 - never let recovery abort a read
                    logger.warning("  page %s: mark recovery failed: %s", index, exc)
                    stacked = {}

                pages.append(
                    PageContent(
                        page_number=index,
                        text=text,
                        tables=tables,
                        image=image,
                        stacked_marks=stacked,
                    )
                )
        return pages

    @staticmethod
    def _page_to_image(page: pdfplumber.page.Page) -> np.ndarray | None:
        try:
            page_image = page.to_image(resolution=200)
            pil_image = page_image.original
            return np.array(pil_image.convert("RGB"))
        except Exception:
            return None
