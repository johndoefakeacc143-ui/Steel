"""Extract text and images from steel structure PDF drawings."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pdfplumber


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
            for index, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                tables = page.extract_tables() or []
                image = self._page_to_image(page)
                try:
                    stacked = recover_stacked_marks(page.chars, text)
                except Exception:  # noqa: BLE001 - never let recovery abort a read
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
