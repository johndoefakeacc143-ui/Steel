"""OCR extraction for image-based steel drawings."""

from __future__ import annotations

import pytesseract
from PIL import Image

from src.image_processor import ImageProcessor
from src.tesseract_setup import ensure_tesseract


class OCREngine:
    """Run Tesseract OCR on preprocessed drawing images."""

    def __init__(self, lang: str = "eng") -> None:
        self.lang = lang
        self.processor = ImageProcessor()
        ensure_tesseract()

    def extract_text(self, image) -> str:
        enhanced = self.processor.enhance_contrast(image)
        processed = self.processor.preprocess(enhanced)
        pil_image = Image.fromarray(processed)
        config = "--psm 6 -c preserve_interword_spaces=1"
        return pytesseract.image_to_string(pil_image, lang=self.lang, config=config)
