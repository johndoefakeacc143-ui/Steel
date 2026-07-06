"""Locate the Tesseract binary automatically (esp. on Windows).

On Windows, Tesseract is usually installed to ``C:\\Program Files\\Tesseract-OCR``
but that directory is not added to ``PATH``, so ``pytesseract`` cannot find it even
though it is installed. This module points ``pytesseract`` at common install
locations so member-length OCR works out of the box.
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

_configured = False

# Common Windows install locations (and a couple of *nix ones for completeness).
_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
    "/opt/homebrew/bin/tesseract",
    "/usr/local/bin/tesseract",
    "/usr/bin/tesseract",
]


def ensure_tesseract() -> bool:
    """Make sure pytesseract can call Tesseract; return True if it can.

    Tries the current configuration first, then ``PATH``, then common install
    paths, setting ``pytesseract.pytesseract.tesseract_cmd`` when a working
    binary is found. Safe to call repeatedly.
    """
    global _configured
    try:
        import pytesseract
    except Exception:  # noqa: BLE001
        return False

    def works() -> bool:
        try:
            pytesseract.get_tesseract_version()
            return True
        except Exception:  # noqa: BLE001
            return False

    if _configured and works():
        return True

    if works():  # already on PATH / already configured
        _configured = True
        return True

    on_path = shutil.which("tesseract")
    candidates = ([on_path] if on_path else []) + _CANDIDATES
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = candidate
            if works():
                _configured = True
                return True

    return False
