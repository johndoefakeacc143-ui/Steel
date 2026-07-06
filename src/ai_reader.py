"""AI vision engine: read steel drawings (PDF or image) with a vision LLM.

Unlike OCR, a multimodal model reads a drawing holistically - faint, rotated, and
graphic (non-selectable) text, plus printed schedule/quantity tables - and returns
structured members. This is the most robust engine for scanned drawings, image
files, and drawings whose BOM lives in graphic tables.

Provider-agnostic (OpenAI-compatible Chat Completions). Configure via environment
(no key stored in the repo):

    ANU_AI_API_KEY  (or OPENAI_API_KEY)   - required
    ANU_AI_BASE_URL (or OPENAI_BASE_URL)  - default https://api.openai.com/v1
    ANU_AI_MODEL                          - default gpt-4o-mini (vision capable)
"""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import urllib.error
import urllib.request
from pathlib import Path

from src.bom_parser import BOMItem
from src.sections import compute_weight_kg

logger = logging.getLogger(__name__)

DEFAULT_BASE_URL = "https://api.openai.com/v1"
DEFAULT_MODEL = "gpt-4o-mini"

PROMPT = (
    "You are an expert steel detailer reading a structural steel drawing. "
    "Identify every steel member and every row of any schedule/quantity table. "
    "Read dimension lines carefully, including faint, low-contrast, and rotated "
    "(vertical) text. Return ONLY JSON of the form {\"members\": [ ... ]} where "
    "each member has keys: mark (string, e.g. B7, C1, BR4), description (string, "
    "e.g. section like ISMB300 / WPB600X300X128.79 or a member type), length_mm "
    "(number in millimetres; convert metres to mm), quantity (integer), grade "
    "(string), weight_kg (number). Use null for anything not determinable. Do not "
    "invent values."
)


class AIConfigError(RuntimeError):
    """Raised when the AI engine is not configured (missing API key)."""


class VisionAIReader:
    """Read a drawing into BOM items using a vision LLM."""

    def __init__(self) -> None:
        self.api_key = os.environ.get("ANU_AI_API_KEY") or os.environ.get("OPENAI_API_KEY")
        self.base_url = (
            os.environ.get("ANU_AI_BASE_URL")
            or os.environ.get("OPENAI_BASE_URL")
            or DEFAULT_BASE_URL
        ).rstrip("/")
        self.model = os.environ.get("ANU_AI_MODEL", DEFAULT_MODEL)

    def available(self) -> bool:
        return bool(self.api_key)

    def read(self, path: Path) -> list[BOMItem]:
        if not self.available():
            raise AIConfigError(
                "AI engine needs an API key. Set ANU_AI_API_KEY (or OPENAI_API_KEY)."
            )
        items: list[BOMItem] = []
        for page_number, image_b64 in enumerate(self._render(path), start=1):
            logger.info("AI reading page %s of %s", page_number, path.name)
            content = self._call_model(image_b64)
            items.extend(self._parse_items(content, page_number))
        return items

    # -- Rendering (PDF pages or an image file) -----------------------------
    def _render(self, path: Path) -> list[str]:
        import cv2
        import numpy as np

        images: list = []
        if path.suffix.lower() == ".pdf":
            import pdfplumber

            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    images.append(np.array(page.to_image(resolution=200).original.convert("RGB")))
        else:
            from PIL import Image

            with Image.open(path) as img:
                images.append(np.array(img.convert("RGB")))

        encoded: list[str] = []
        for rgb in images:
            enhanced = self._enhance(rgb, cv2)
            ok, buf = cv2.imencode(".png", cv2.cvtColor(enhanced, cv2.COLOR_RGB2BGR))
            if ok:
                encoded.append(base64.b64encode(buf.tobytes()).decode("ascii"))
        return encoded

    @staticmethod
    def _enhance(rgb, cv2):
        lab = cv2.cvtColor(rgb, cv2.COLOR_RGB2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        return cv2.cvtColor(cv2.merge((clahe.apply(l), a, b)), cv2.COLOR_LAB2RGB)

    # -- Model call ---------------------------------------------------------
    def _call_model(self, image_b64: str) -> str:
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": PROMPT},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
                    ],
                }
            ],
            "temperature": 0,
            "max_tokens": 4000,
        }
        request = urllib.request.Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=180) as response:
                body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f"AI API error {exc.code}: {exc.read().decode('utf-8', 'ignore')}") from exc
        return body["choices"][0]["message"]["content"]

    # -- Parsing ------------------------------------------------------------
    def _parse_items(self, content: str, page_number: int) -> list[BOMItem]:
        data = self._extract_json(content)
        rows = data.get("members", data) if isinstance(data, dict) else data
        if not isinstance(rows, list):
            logger.warning("AI response had no member list; got: %.200s", content)
            return []

        items: list[BOMItem] = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            mark = self._str(row.get("mark")) or self._str(row.get("member_id"))
            description = self._str(row.get("description")) or self._str(row.get("section_size")) or ""
            length = self._num(row.get("length_mm"))
            weight = self._num(row.get("weight_kg"))
            if weight is None and description:
                weight = compute_weight_kg(description, length)
            if not mark and not description:
                continue
            items.append(
                BOMItem(
                    mark=mark or "NA",
                    description=description,
                    quantity=self._int(row.get("quantity")) or 1,
                    length=length,
                    grade=self._str(row.get("grade")) or "",
                    weight=weight,
                    source_page=page_number,
                )
            )
        return items

    @staticmethod
    def _extract_json(content: str):
        text = re.sub(r"^```(?:json)?|```$", "", content.strip(), flags=re.MULTILINE).strip()
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            match = re.search(r"[\[{].*[\]}]", text, re.DOTALL)
            if match:
                return json.loads(match.group(0))
            raise

    @staticmethod
    def _str(value):
        if value is None:
            return None
        text = str(value).strip()
        return text or None

    @staticmethod
    def _num(value):
        try:
            return float(value) if value is not None and str(value).strip() != "" else None
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _int(value):
        try:
            return int(float(value)) if value is not None and str(value).strip() != "" else None
        except (TypeError, ValueError):
            return None
