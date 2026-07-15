"""OpenCV preprocessing tuned for steel engineering drawings / blueprints."""

from __future__ import annotations

import cv2
import numpy as np
from PIL import Image


def pil_to_bgr(image: Image.Image) -> np.ndarray:
    """Convert a PIL image to OpenCV BGR."""
    rgb = np.array(image.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def preprocess_for_ocr(bgr: np.ndarray, upscale: float = 1.6) -> np.ndarray:
    """
    Denoise and binarize blueprint-style scans for Tesseract.

    Upscales thin linework, applies NL-means denoise, then adaptive threshold
    so pale pencil/plotter text becomes high-contrast OCR input.
    """
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    if upscale and upscale != 1.0:
        gray = cv2.resize(
            gray,
            None,
            fx=upscale,
            fy=upscale,
            interpolation=cv2.INTER_CUBIC,
        )
    denoise = cv2.fastNlMeansDenoising(gray, h=10)
    binary = cv2.adaptiveThreshold(
        denoise,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        11,
    )
    return binary


def preprocess_variants(bgr: np.ndarray) -> list[np.ndarray]:
    """
    Produce a few OCR-friendly variants.

    Some drawings OCR better with adaptive binary; others with Otsu or
    inverted blueprints (white lines on dark). Callers can OCR each and merge.
    """
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    denoise = cv2.fastNlMeansDenoising(gray, h=10)

    adaptive = cv2.adaptiveThreshold(
        denoise, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 11
    )
    _, otsu = cv2.threshold(denoise, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    inverted = cv2.bitwise_not(adaptive)

    # Prefer bright-background variants first
    mean_val = float(np.mean(denoise))
    if mean_val < 127:
        return [inverted, adaptive, otsu, denoise]
    return [adaptive, otsu, denoise]
