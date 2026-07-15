"""Spatially associate plan dimension text with nearby member marks via OCR."""

from __future__ import annotations

import collections
import logging
import math
import re
from typing import Any

import cv2
import numpy as np
import pytesseract
from PIL import Image

logger = logging.getLogger("steel_ocr.spatial")

BEAM_RE = re.compile(r"^B[1-9]\d*$")
BRACE_RE = re.compile(r"^BR\d+$")
BP_RE = re.compile(r"^BP\d+$")
PB_RE = re.compile(r"^PB\d+$")
PLAT_RE = re.compile(r"^PB[A-G]$|^PAB$")
DIM_RE = re.compile(r"^(\d{3,5})$")

OCR_FIXES = {
    "BRI": "BR1",
    "BA": "B4",
    "BG": "B6",
    "BS": "B5",
    "BE": "B8",
    "BQ": "B6",
}


def _normalize_mark(raw: str) -> str:
    t = re.sub(r"[^A-Za-z0-9]", "", raw).upper()
    t = OCR_FIXES.get(t, t)
    if t.endswith("L") and BEAM_RE.fullmatch(t[:-1]):
        t = t[:-1]
    return t


def _dedupe(items: list[dict[str, Any]], tol: float = 25.0) -> list[dict[str, Any]]:
    kept: list[dict[str, Any]] = []
    for it in sorted(items, key=lambda x: (x["text"], x["cy"], x["cx"])):
        if any(
            abs(it["cx"] - k["cx"]) < tol
            and abs(it["cy"] - k["cy"]) < tol
            and it["text"] == k["text"]
            for k in kept
        ):
            continue
        kept.append(it)
    return kept


def ocr_tokens_with_boxes(
    image: Image.Image,
    *,
    upscale: float = 1.25,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    OCR a plan image and return (marks, dimensions) with pixel centers.

    Dimensions are bare metric lengths typically written along members (900–30000).
    """
    bgr = cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    variants: list[tuple[str, np.ndarray, float]] = [
        ("gray", gray, 1.0),
        (
            "bin",
            cv2.adaptiveThreshold(
                cv2.fastNlMeansDenoising(
                    cv2.resize(gray, None, fx=upscale, fy=upscale, interpolation=cv2.INTER_CUBIC),
                    h=7,
                ),
                255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY,
                31,
                9,
            ),
            upscale,
        ),
    ]

    marks: list[dict[str, Any]] = []
    dims: list[dict[str, Any]] = []

    for _name, arr, scale in variants:
        try:
            df = pytesseract.image_to_data(
                arr, config="--oem 3 --psm 11", output_type=pytesseract.Output.DATAFRAME
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("spatial OCR failed: %s", exc)
            continue
        df = df.dropna(subset=["text"])
        df["text"] = df["text"].astype(str).str.strip()
        df = df[(df["text"] != "") & (df["conf"].astype(float) > 15)]
        for r in df.itertuples():
            raw = str(r.text).strip()
            t = _normalize_mark(raw)
            cx = (float(r.left) + float(r.width) / 2) / scale
            cy = (float(r.top) + float(r.height) / 2) / scale
            item = {"text": t, "cx": cx, "cy": cy}
            if (
                BEAM_RE.fullmatch(t)
                or BRACE_RE.fullmatch(t)
                or BP_RE.fullmatch(t)
                or PB_RE.fullmatch(t)
                or PLAT_RE.fullmatch(t)
            ):
                marks.append(item)
            m = DIM_RE.fullmatch(raw)
            if m:
                val = int(m.group(1))
                if 900 <= val <= 30000:
                    dims.append({"text": str(val), "cx": cx, "cy": cy})

    return _dedupe(marks, 30), _dedupe(dims, 20)


def associate_dim_to_mark(mark: dict[str, Any], dims: list[dict[str, Any]]) -> str:
    """
    Pick the best nearby plan dimension for a member mark.

    Prefers dimensions on the same horizontal/vertical band (including dims
    directly above/beside the mark). When several dims align — e.g. bay split
    3000 and overall span 6000 — prefers the larger span (member length).
    """
    horiz: list[tuple[float, int, str]] = []
    vert: list[tuple[float, int, str]] = []
    near: list[tuple[float, int, str]] = []

    for d in dims:
        dx = abs(d["cx"] - mark["cx"])
        dy = abs(d["cy"] - mark["cy"])
        dxy = math.hypot(dx, dy)
        val = int(d["text"])
        # Same horizontal band (dim above/below beam) — allow dx≈0 (directly above)
        if dy < 120 and dx < 550:
            # Prefer closer alignment, but keep larger values competitive
            horiz.append((dy * 2.0 + dx * 0.15, -val, d["text"]))
        # Same vertical band (dim left/right of vertical member)
        elif dx < 120 and dy < 550:
            vert.append((dx * 2.0 + dy * 0.15, -val, d["text"]))
        elif dxy < 200:
            near.append((dxy, -val, d["text"]))

    def pick(group: list[tuple[float, int, str]], slack: float = 140.0) -> str:
        if not group:
            return ""
        group = sorted(group)
        best = group[0][0]
        pool = [g for g in group if g[0] <= best + slack]
        # Largest dimension among well-aligned candidates (6000 over 3000)
        pool.sort(key=lambda g: (g[1], g[0]))
        return pool[0][2]

    return pick(horiz) or pick(vert) or pick(near) or ""


def infer_sizes_from_plan_image(image: Image.Image) -> dict[str, str]:
    """
    Return {member_name: size} by OCR'ing plan dimensions near marks.

    Size is the mode of associated dimensions across all instances of that mark.
    """
    marks, dims = ocr_tokens_with_boxes(image)
    logger.info("spatial OCR: %s marks, %s dims", len(marks), len(dims))
    if not dims:
        return {}

    by_mark: dict[str, list[str]] = collections.defaultdict(list)
    for m in marks:
        if not (
            BEAM_RE.fullmatch(m["text"])
            or BRACE_RE.fullmatch(m["text"])
            or PB_RE.fullmatch(m["text"])
        ):
            continue
        size = associate_dim_to_mark(m, dims)
        if size:
            by_mark[m["text"]].append(size)

    inferred: dict[str, str] = {}
    for mark, vals in by_mark.items():
        ctr = collections.Counter(vals)
        # Prefer the largest frequently-seen span (overall 6000 over bay half 3000)
        ranked = ctr.most_common()
        mode_val, mode_n = ranked[0]
        best = mode_val
        for val, n in ranked:
            if n >= max(1, mode_n // 2) and int(val) > int(best):
                best = val
        inferred[mark] = best
    return inferred


def apply_spatial_sizes(
    members: list[dict[str, Any]],
    sizes: dict[str, str],
    name_key: str = "Member Name",
    size_key: str = "Size",
) -> list[dict[str, Any]]:
    """Fill empty Size fields from spatially inferred plan dimensions."""
    if not sizes:
        return members
    for row in members:
        name = str(row.get(name_key) or "").upper()
        if not row.get(size_key) and name in sizes:
            row[size_key] = sizes[name]
    return members
