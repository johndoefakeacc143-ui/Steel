#!/usr/bin/env python3
"""Diagnose why member lengths are not being captured.

Run:  python tools/diagnose.py input/ABC.pdf
Paste the full output when reporting an issue.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def line(msg: str) -> None:
    print(msg, flush=True)


def main() -> int:
    pdf = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("input/ABC.pdf")
    line(f"# Diagnosing length capture on: {pdf}")

    # 1. Package versions
    line("\n[1] Package versions")
    for name in ("pdfplumber", "pypdfium2", "pytesseract", "cv2", "numpy", "PIL"):
        try:
            mod = __import__(name)
            line(f"    {name}: {getattr(mod, '__version__', 'unknown')}")
        except Exception as exc:  # noqa: BLE001
            line(f"    {name}: NOT IMPORTABLE ({exc})")

    # 2. Tesseract binary
    line("\n[2] Tesseract binary")
    try:
        import pytesseract

        line(f"    tesseract_cmd = {pytesseract.pytesseract.tesseract_cmd}")
        line(f"    version = {pytesseract.get_tesseract_version()}")
    except Exception as exc:  # noqa: BLE001
        line(f"    ERROR: {exc}")
        line("    -> pytesseract cannot run Tesseract. On Windows set: "
             "pytesseract.pytesseract.tesseract_cmd = r'C:\\\\Program Files\\\\Tesseract-OCR\\\\tesseract.exe'")

    # 3. Page rendering (needed for all OCR)
    line("\n[3] PDF page rendering (page.to_image)")
    image = None
    try:
        import numpy as np
        import pdfplumber

        with pdfplumber.open(pdf) as doc:
            page = doc.pages[0]
            line(f"    page text length = {len(page.extract_text() or '')}")
            rendered = page.to_image(resolution=150)
            image = np.array(rendered.original.convert("RGB"))
            line(f"    rendered image shape = {image.shape}  <-- rendering OK")
    except Exception as exc:  # noqa: BLE001
        line(f"    RENDER FAILED: {exc}")
        line("    -> Without rendering, OCR cannot read dimensions. Ensure pypdfium2 is installed: "
             "pip install pypdfium2")

    # 4. OCR sample
    if image is not None:
        line("\n[4] OCR sample on rendered page")
        try:
            import cv2
            import pytesseract

            gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
            txt = pytesseract.image_to_string(gray, config="--psm 11")
            toks = [t for t in txt.split() if t.strip()]
            line(f"    OCR produced {len(toks)} tokens; sample: {toks[:15]}")
        except Exception as exc:  # noqa: BLE001
            line(f"    OCR FAILED: {exc}")

    # 5. Full dimension estimation
    line("\n[5] Dimension length estimation")
    try:
        from src.dimension_ocr import DimensionEstimator
        from src.pdf_reader import PDFReader

        pages = PDFReader(pdf).read()
        marks, defaults = DimensionEstimator().estimate(pages)
        line(f"    per-mark lengths = {marks}")
        line(f"    page default lengths = {defaults}")
        if not marks and not defaults:
            line("    -> No lengths estimated. See stages [2]/[3]/[4] above for the cause.")
        else:
            line("    -> Lengths estimated OK; they will appear in the BOM.")
    except Exception as exc:  # noqa: BLE001
        line(f"    ESTIMATION FAILED: {exc}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
