"""Estimate member lengths from drawing dimension lines via OCR.

General-arrangement drawings rarely tabulate member lengths; instead the length
is shown as a dimension annotation next to the member. Those dimension labels are
drawn as graphics (often rotated 90 degrees) and are not part of the PDF text
layer, so they are only recoverable through OCR of the rendered page image.

For each page this module:
  1. OCRs the page image (horizontal marks + horizontal dimensions).
  2. OCRs a 90-degrees-rotated copy to recover vertical dimensions, mapping their
     coordinates back to the original image.
  3. Assigns every detected member mark the nearest dimension value, then takes
     the most common value per mark as its representative length.

Only linear members (beams ``B2``..``B9`` and bracings ``BR#``) receive a length;
base plates and plan-bay references do not have a meaningful length.
"""

from __future__ import annotations

import logging
import re
from collections import Counter, defaultdict

from src.bom_parser import DRAWING_MARK_PATTERN

logger = logging.getLogger(__name__)

# Cap the image long side before OCR (px). Larger renders are downscaled to keep
# OCR fast on big multi-page drawing sets.
_MAX_OCR_LONG_SIDE = 4500

_MARK_RE = re.compile(r"^(B[2-9]|BR\d+|BP[12]|PB[1-9A-Z])$", re.IGNORECASE)
_LINEAR_MARK_RE = re.compile(r"^(B[2-9]|BR\d+)$", re.IGNORECASE)
_DIM_RE = re.compile(r"^\d{3,5}$")

# Plausible member dimension range in millimetres (filters OCR noise).
_MIN_DIM_MM = 100
_MAX_DIM_MM = 30000
# A mark must have at least this fraction of its votes agree to be trusted.
_MIN_MODE_FRACTION = 0.4


class DimensionEstimator:
    """Estimate representative member lengths (mm) from page images via OCR."""

    def estimate(
        self, pages: list
    ) -> tuple[dict[tuple[int, str], float], dict[int, float]]:
        """Return (per-mark lengths, per-page default beam length).

        - per-mark: ``{(page_number, MARK): length_mm}`` for confidently
          estimated linear members.
        - per-page default: ``{page_number: length_mm}`` = the dominant beam
          dimension on the page, used as a fallback for beam marks OCR could not
          individually resolve.
        """
        try:
            import cv2  # noqa: F401
            import numpy as np  # noqa: F401
            import pytesseract  # noqa: F401

            from src.tesseract_setup import ensure_tesseract

            ensure_tesseract()
        except Exception as exc:  # noqa: BLE001
            logger.warning("OCR dependencies unavailable (%s); skipping dimensions", exc)
            return {}, {}

        mark_lengths: dict[tuple[int, str], float] = {}
        page_defaults: dict[int, float] = {}

        total = len(pages)
        for index, page in enumerate(pages, start=1):
            image = getattr(page, "image", None)
            if image is None:
                continue
            # Skip the expensive OCR on pages that have no member marks - there is
            # nothing to attach a dimension to (big speed-up on multi-page sets).
            if not self._page_has_marks(page):
                logger.info(
                    "Page %s/%s: no member marks - skipping dimension OCR", index, total
                )
                continue
            logger.info(
                "Page %s/%s: reading dimensions via OCR (image %sx%s)...",
                index,
                total,
                image.shape[1],
                image.shape[0],
            )
            try:
                per_mark, default = self._estimate_page(image)
            except Exception as exc:  # noqa: BLE001 - OCR must never abort a run
                logger.warning("Dimension OCR failed on page %s: %s", page.page_number, exc)
                continue
            for mark, length in per_mark.items():
                mark_lengths[(page.page_number, mark)] = length
            if default is not None:
                page_defaults[page.page_number] = default
        return mark_lengths, page_defaults

    @staticmethod
    def _page_has_marks(page) -> bool:
        if getattr(page, "stacked_marks", None):
            return True
        return bool(DRAWING_MARK_PATTERN.search(getattr(page, "text", "") or ""))

    def _estimate_page(self, image):
        import cv2
        import numpy as np
        import pytesseract

        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        # Downscale very large renders before OCR: dimension text stays legible but
        # OCR time drops roughly with the pixel count. Association is relative, so
        # a uniform scale does not affect nearest-dimension matching.
        long_side = max(gray.shape[:2])
        if long_side > _MAX_OCR_LONG_SIDE:
            scale = _MAX_OCR_LONG_SIDE / long_side
            gray = cv2.resize(
                gray, (int(gray.shape[1] * scale), int(gray.shape[0] * scale)),
                interpolation=cv2.INTER_AREA,
            )
        height = gray.shape[0]

        marks: list[tuple[str, float, float]] = []
        dims: list[tuple[int, float, float]] = []

        logger.info("    horizontal OCR pass...")
        for text, cx, cy in self._ocr_tokens(gray, pytesseract):
            if _MARK_RE.match(text):
                marks.append((text.upper(), cx, cy))
            elif _DIM_RE.match(text):
                value = int(text)
                if _MIN_DIM_MM <= value <= _MAX_DIM_MM:
                    dims.append((value, cx, cy))

        # Rotated pass recovers vertical dimension labels.
        logger.info("    vertical OCR pass...")
        rotated = cv2.rotate(gray, cv2.ROTATE_90_CLOCKWISE)
        for text, cx, cy in self._ocr_tokens(rotated, pytesseract):
            if _DIM_RE.match(text):
                value = int(text)
                if _MIN_DIM_MM <= value <= _MAX_DIM_MM:
                    dims.append((value, cy, height - 1 - cx))

        if not marks or not dims:
            return {}, None

        dim_xy = np.array([[d[1], d[2]] for d in dims], dtype=float)
        dim_val = [d[0] for d in dims]

        votes: dict[str, list[int]] = defaultdict(list)
        beam_votes: list[int] = []
        for mark, mx, my in marks:
            distances = np.hypot(dim_xy[:, 0] - mx, dim_xy[:, 1] - my)
            nearest = dim_val[int(np.argmin(distances))]
            votes[mark].append(nearest)
            if _LINEAR_MARK_RE.match(mark):
                beam_votes.append(nearest)

        per_mark: dict[str, float] = {}
        for mark, values in votes.items():
            if not _LINEAR_MARK_RE.match(mark):
                continue
            value, count = Counter(values).most_common(1)[0]
            if count / len(values) >= _MIN_MODE_FRACTION:
                per_mark[mark] = float(value)

        default = float(Counter(beam_votes).most_common(1)[0][0]) if beam_votes else None
        return per_mark, default
    @staticmethod
    def _ocr_tokens(image, pytesseract):
        data = pytesseract.image_to_data(
            image, config="--psm 11", output_type=pytesseract.Output.DICT
        )
        tokens = []
        for i, text in enumerate(data["text"]):
            token = text.strip()
            if not token:
                continue
            cx = data["left"][i] + data["width"][i] / 2.0
            cy = data["top"][i] + data["height"][i] / 2.0
            tokens.append((token, cx, cy))
        return tokens
