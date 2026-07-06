"""AI vision reader: extract steel members from drawings using a vision LLM.

Unlike OCR (Tesseract/EasyOCR), a multimodal model reads a drawing holistically -
including faint, low-contrast, and rotated dimension text - and returns structured
member data directly. This is the recommended engine for General Arrangement
drawings whose dimensions are drawn as graphics rather than selectable text.

Provider-agnostic: it speaks the OpenAI-compatible Chat Completions API, so it
works with OpenAI, OpenRouter, Azure, or a local server (vLLM/Ollama/llama.cpp)
by setting the base URL. Configure via environment variables:

    ANU_AI_API_KEY   (or OPENAI_API_KEY)          - required
    ANU_AI_BASE_URL  (or OPENAI_BASE_URL)         - default https://api.openai.com/v1
    ANU_AI_MODEL                                   - default gpt-4o-mini (vision capable)

No API key is stored in the repo; the key is read from the environment at runtime.
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

from src.extractor import Member
from src.sections import compute_weight_kg

logger = logging.getLogger(__name__)

DEFAULT_BASE_URL = "https://api.openai.com/v1"
DEFAULT_MODEL = "gpt-4o-mini"

PROMPT = (
    "You are an expert steel detailer reading a structural steel drawing. "
    "Identify every steel member in the image. Read dimension lines carefully, "
    "including faint, low-contrast, and rotated (vertical) text, to determine "
    "lengths. Return ONLY a JSON object of the form "
    '{"members": [ ... ]} where each member has these keys: '
    "member_id (string), member_type (one of Beam, Column, Bracing, Plate, "
    "Girder), section_size (e.g. ISMB 300, ISMC 150, RHS 100x100x4, PL 200x10), "
    "length_mm (number in millimetres; convert metres to mm), quantity (integer), "
    "elevation_m (number), grid_location (string), weight_kg (number). "
    "Use null for anything you cannot determine. Do not invent values."
)


class AIConfigError(RuntimeError):
    """Raised when the AI engine is not configured (e.g. missing API key)."""


class VisionAIReader:
    """Read steel members from PDF drawings with a vision LLM."""

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

    def read(self, pdf_path: Path) -> list[Member]:
        if not self.available():
            raise AIConfigError(
                "AI engine needs an API key. Set ANU_AI_API_KEY (or OPENAI_API_KEY)."
            )
        members: list[Member] = []
        for page_index, image_b64 in enumerate(self._render_pages(pdf_path), start=1):
            logger.info("AI reading page %s of %s", page_index, pdf_path.name)
            content = self._call_model(image_b64)
            members.extend(self._parse_members(content, pdf_path.name))
        return members

    # -- Rendering ----------------------------------------------------------
    def _render_pages(self, pdf_path: Path) -> list[str]:
        import cv2
        import numpy as np
        import pdfplumber

        images: list[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                rgb = np.array(page.to_image(resolution=200).original.convert("RGB"))
                enhanced = self._enhance(rgb, cv2)
                ok, buf = cv2.imencode(".png", cv2.cvtColor(enhanced, cv2.COLOR_RGB2BGR))
                if ok:
                    images.append(base64.b64encode(buf.tobytes()).decode("ascii"))
        return images

    @staticmethod
    def _enhance(rgb, cv2):
        """Boost contrast so faint dimension text is legible to the model."""
        lab = cv2.cvtColor(rgb, cv2.COLOR_RGB2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        merged = cv2.merge((clahe.apply(l), a, b))
        return cv2.cvtColor(merged, cv2.COLOR_LAB2RGB)

    # -- Model call ---------------------------------------------------------
    def _call_model(self, image_b64: str) -> str:
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": PROMPT},
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                        },
                    ],
                }
            ],
            "temperature": 0,
            "max_tokens": 4000,
        }
        request = urllib.request.Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "ignore")
            raise RuntimeError(f"AI API error {exc.code}: {detail}") from exc
        return body["choices"][0]["message"]["content"]

    # -- Parsing ------------------------------------------------------------
    def _parse_members(self, content: str, source: str) -> list[Member]:
        data = self._extract_json(content)
        rows = data.get("members", data) if isinstance(data, dict) else data
        if not isinstance(rows, list):
            logger.warning("AI response had no member list; got: %.200s", content)
            return []

        members: list[Member] = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            member = Member(
                member_id=self._str(row.get("member_id")),
                member_type=self._str(row.get("member_type")),
                section_size=self._str(row.get("section_size")),
                length_mm=self._num(row.get("length_mm")),
                quantity=self._int(row.get("quantity")),
                elevation_m=self._num(row.get("elevation_m")),
                grid_location=self._str(row.get("grid_location")),
                weight_kg=self._num(row.get("weight_kg")),
                source_file=source,
            )
            if member.quantity is None:
                member.quantity = 1
            if member.weight_kg is None and member.section_size:
                member.weight_kg = compute_weight_kg(member.section_size, member.length_mm)
            if member.member_id or member.section_size:
                members.append(member)
        return members

    @staticmethod
    def _extract_json(content: str):
        text = content.strip()
        text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
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
