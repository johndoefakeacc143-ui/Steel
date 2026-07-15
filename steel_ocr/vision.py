"""
Optional vision-LLM assist — reads a drawing like a human detailer.

Used only when OPENAI_API_KEY or GEMINI_API_KEY is set. Never hardcodes
members; every call looks at the uploaded page image(s) for this run.
"""

from __future__ import annotations

import base64
import collections
import io
import json
import logging
import os
import re
from pathlib import Path
from typing import Any

from PIL import Image

logger = logging.getLogger("steel_ocr.vision")

VISION_PROMPT = """You are a structural steel detailer reading an engineering GA / framing plan.

Look at the drawing carefully (as a human would) and list EVERY steel member mark you can see.

For each distinct mark + length combination provide:
- member_name: mark as printed (B1, B4, B6, B7, BR1, BR4, BP1, C1, PB1, PBA, …)
- member_type: beam | column | brace | base_plate | other
- quantity: how many times this exact mark appears with this size (count labels on the sheet)
- size: length or section in mm if readable from nearby dimension text
- length_note: for DIAGONAL braces only, write the formula √(a² + b²) using the
  horizontal run and vertical rise next to that brace (e.g. √(3000² + 1500²)).
  Leave empty for straight beams/columns.
- notes: brief placement note if useful

Rules:
- Same mark at different lengths = SEPARATE items (e.g. B6@1000 and B6@2000).
- Do NOT invent marks that are not on the drawing.
- Prefer dimensions written next to the member over title-block scales.
- Return JSON only: {"items":[...]}
"""


def _api_keys() -> dict[str, str]:
    def clean(raw: str | None) -> str:
        if not raw:
            return ""
        key = raw.strip().strip('"').strip("'")
        placeholders = {
            "",
            "sk-...",
            "your_key_here",
            "your_openai_api_key_here",
            "your_google_gemini_api_key_here",
        }
        if key.lower() in placeholders or key.lower().startswith("sk-your"):
            return ""
        return key

    return {
        "openai": clean(os.getenv("OPENAI_API_KEY")),
        "gemini": clean(os.getenv("GEMINI_API_KEY")),
    }


def vision_available() -> bool:
    keys = _api_keys()
    return bool(keys["openai"] or keys["gemini"])


def _image_to_png_b64(image: Image.Image, max_side: int = 2000) -> str:
    im = image.convert("RGB")
    w, h = im.size
    scale = min(1.0, max_side / max(w, h))
    if scale < 1.0:
        im = im.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _parse_items(content: str) -> list[dict[str, Any]]:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", text)
        if not m:
            return []
        try:
            data = json.loads(m.group(0))
        except json.JSONDecodeError:
            return []
    items = data.get("items") if isinstance(data, dict) else data
    if not isinstance(items, list):
        return []
    return [i for i in items if isinstance(i, dict)]


def _call_openai(images: list[Image.Image], api_key: str) -> list[dict[str, Any]]:
    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    model = os.getenv("OPENAI_VISION_MODEL", "gpt-4o-mini")
    content: list[dict[str, Any]] = [{"type": "text", "text": VISION_PROMPT}]
    for im in images[:4]:
        b64 = _image_to_png_b64(im)
        content.append(
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"},
            }
        )
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": content}],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    return _parse_items(resp.choices[0].message.content or "")


def _call_gemini(images: list[Image.Image], api_key: str) -> list[dict[str, Any]]:
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    model = os.getenv("GEMINI_VISION_MODEL", "gemini-2.5-flash")
    parts: list[Any] = [VISION_PROMPT]
    for im in images[:4]:
        buf = io.BytesIO()
        rgb = im.convert("RGB")
        w, h = rgb.size
        scale = min(1.0, 2000 / max(w, h))
        if scale < 1.0:
            rgb = rgb.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
        rgb.save(buf, format="PNG")
        parts.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/png"))

    resp = client.models.generate_content(
        model=model,
        contents=parts,
        config=types.GenerateContentConfig(
            temperature=0.1,
            response_mime_type="application/json",
        ),
    )
    return _parse_items(resp.text or "")


def vision_extract_members(images: list[Image.Image]) -> list[dict[str, Any]]:
    """
    Ask a vision model to read this upload's page images.

    Returns normalized rows: Member Name, Quantity, Size, Length Note, type.
    Empty list if no API key or on failure — OCR path still runs alone.
    """
    keys = _api_keys()
    if not images:
        return []
    items: list[dict[str, Any]] = []
    try:
        if keys["openai"]:
            logger.info("Vision assist: OpenAI (%s page images)", len(images))
            items = _call_openai(images, keys["openai"])
        elif keys["gemini"]:
            logger.info("Vision assist: Gemini (%s page images)", len(images))
            items = _call_gemini(images, keys["gemini"])
        else:
            logger.info("Vision assist OFF — set OPENAI_API_KEY or GEMINI_API_KEY")
            return []
    except Exception as exc:  # noqa: BLE001
        logger.warning("Vision assist failed: %s", exc)
        return []

    rows: list[dict[str, Any]] = []
    for it in items:
        name = str(it.get("member_name") or it.get("mark") or "").strip().upper()
        name = re.sub(r"\s+", "", name)
        if not name:
            continue
        try:
            qty = int(it.get("quantity") or it.get("qty") or 1)
        except (TypeError, ValueError):
            qty = 1
        qty = max(1, qty)
        size = str(it.get("size") or it.get("length") or "").strip()
        size = re.sub(r"[^\d.]", "", size.split()[0]) if size else ""
        note = str(it.get("length_note") or "").strip()
        mtype = str(it.get("member_type") or "beam").strip().lower()
        rows.append(
            {
                "Member Name": name,
                "Quantity": qty,
                "Size": size,
                "Length": size,
                "Length Note": note,
                "member_type": mtype,
                "Page": 1,
                "Source": "vision",
            }
        )
    logger.info("Vision returned %s member rows", len(rows))
    return rows


def merge_ocr_and_vision(
    ocr_rows: list[dict[str, Any]],
    vision_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Merge OCR takeoff with vision assist.

    - Keep OCR rows as the base (counts from text layer + OCR are preferred)
    - Add vision-only marks (e.g. BR4 not in PDF text) when OCR missed them
    - Fill blank OCR sizes from matching vision name when vision has a size
    """
    if not vision_rows:
        return ocr_rows

    def key(row: dict[str, Any]) -> tuple[str, str]:
        return (
            str(row.get("Member Name") or "").upper(),
            str(row.get("Size") or ""),
        )

    by_key: dict[tuple[str, str], dict[str, Any]] = {key(r): dict(r) for r in ocr_rows}
    ocr_names = {str(r.get("Member Name") or "").upper() for r in ocr_rows}

    # Fill blank sizes from vision (same mark name)
    vision_by_name: dict[str, list[dict[str, Any]]] = collections.defaultdict(list)
    for vr in vision_rows:
        vision_by_name[str(vr["Member Name"]).upper()].append(vr)

    for k, row in list(by_key.items()):
        if row.get("Size"):
            continue
        name = k[0]
        cands = vision_by_name.get(name) or []
        if len(cands) == 1 and cands[0].get("Size"):
            row["Size"] = cands[0]["Size"]
            row["Length"] = cands[0].get("Length") or cands[0]["Size"]
            if cands[0].get("Length Note"):
                row["Length Note"] = cands[0]["Length Note"]

    # Add vision marks OCR completely missed
    for vr in vision_rows:
        name = str(vr["Member Name"]).upper()
        if name in ocr_names:
            # If OCR has the mark but not this size variant, add the variant
            k = key(vr)
            if k not in by_key and vr.get("Size"):
                by_key[k] = dict(vr)
            continue
        by_key[key(vr)] = dict(vr)

    return list(by_key.values())


def classify_vision_buckets(
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """Split vision/OCR rows into beams, columns, base plates."""
    beams: list[dict[str, Any]] = []
    columns: list[dict[str, Any]] = []
    plates: list[dict[str, Any]] = []
    for row in rows:
        name = str(row.get("Member Name") or "").upper()
        mtype = str(row.get("member_type") or "").lower()
        if mtype == "base_plate" or name.startswith("BP"):
            plates.append(row)
        elif mtype == "column" or name.startswith("COL") or (
            name.startswith("C") and not name.startswith("BR") and re.fullmatch(r"C\d+", name)
        ):
            columns.append(row)
        elif name.startswith(("PB", "PA")) and re.fullmatch(r"PB[A-G]|PAB", name):
            columns.append(row)
        else:
            beams.append(row)
    return beams, columns, plates
