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


def diagonal_length_mm(leg_a_mm: float, leg_b_mm: float) -> float:
    """
    Length of a diagonally placed brace / member.

        L = √(a² + b²)
    """
    return math.hypot(float(leg_a_mm), float(leg_b_mm))


def format_mm_length(mm: float) -> str:
    """Format plan length as whole millimetres (engineering takeoff)."""
    return str(int(round(mm)))


def _classify_nearby_dims(
    mark: dict[str, Any], dims: list[dict[str, Any]]
) -> tuple[list[tuple[float, int, str]], list[tuple[float, int, str]], list[tuple[float, int, str]]]:
    """Split nearby dims into horizontal-band, vertical-band, and near groups."""
    horiz: list[tuple[float, int, str]] = []
    vert: list[tuple[float, int, str]] = []
    near: list[tuple[float, int, str]] = []

    for d in dims:
        dx = abs(d["cx"] - mark["cx"])
        dy = abs(d["cy"] - mark["cy"])
        dxy = math.hypot(dx, dy)
        val = int(d["text"])
        # Same horizontal band (dim above/below member)
        if dy < 120 and dx < 550:
            horiz.append((dy * 2.0 + dx * 0.15, -val, d["text"]))
        # Same vertical band (dim left/right of member)
        elif dx < 120 and dy < 550:
            vert.append((dx * 2.0 + dy * 0.15, -val, d["text"]))
        elif dxy < 280:
            near.append((dxy, -val, d["text"]))
    return horiz, vert, near


def _pick_from_group(
    group: list[tuple[float, int, str]],
    *,
    slack: float = 140.0,
    prefer_larger: bool = True,
) -> str:
    if not group:
        return ""
    group = sorted(group)
    best = group[0][0]
    pool = [g for g in group if g[0] <= best + slack]
    if prefer_larger:
        pool.sort(key=lambda g: (g[1], g[0]))  # largest dim among aligned
    return pool[0][2]


def associate_dim_to_mark(mark: dict[str, Any], dims: list[dict[str, Any]]) -> str:
    """
    Pick the best nearby plan dimension for a straight member mark.

    Prefers dimensions on the same horizontal/vertical band. When several dims
    align (e.g. bay 3000 and overall 6000), prefers the larger span.
    """
    horiz, vert, near = _classify_nearby_dims(mark, dims)
    return (
        _pick_from_group(horiz)
        or _pick_from_group(vert)
        or _pick_from_group(near)
        or ""
    )


def pick_diagonal_legs(
    mark: dict[str, Any],
    dims: list[dict[str, Any]],
) -> tuple[float, float] | None:
    """
    For a diagonal brace (BR1…), pick horizontal run (a) and vertical rise (b).

    Returns (a_mm, b_mm) for L = √(a² + b²), or None if both legs cannot be found.

    Prefers the bay half-span (e.g. 3000) over the overall grid span (6000) for
    the horizontal run of a chevron / inverted-V brace.
    """
    horiz_cands: list[tuple[float, float]] = []  # (score, value)
    vert_cands: list[tuple[float, float]] = []

    for d in dims:
        dx = abs(d["cx"] - mark["cx"])
        dy = abs(d["cy"] - mark["cy"])
        val = float(d["text"])
        if val < 900:
            continue
        # Top/bottom dimension string → horizontal run of the brace bay
        is_horiz_band = dy < 150 and dy >= 8 and dx < 650
        # Side story-height dim: beside the mark (same height band, far in x)
        # or clearly lateral vs the top dimension string.
        is_side = (dy < 45 and dx >= 120) or (dx > 220 and dx > dy * 1.5)
        if is_horiz_band and not is_side:
            horiz_cands.append((dy * 2.0 + dx * 0.2, val))
        if is_side and dy < 250 and dx < 900:
            score = dx * 0.4 + dy * 2.0
            if val in (1500.0, 2000.0, 2500.0):
                score *= 0.5
            vert_cands.append((score, val))
        elif not is_horiz_band and not is_side and dx < 200 and dy < 350:
            vert_cands.append((dx * 1.5 + dy * 0.5, val))

    def best_leg(cands: list[tuple[float, float]], *, prefer_smaller_bay: bool) -> float | None:
        if not cands:
            return None
        cands = sorted(cands)
        best_score = cands[0][0]
        pool = [v for s, v in cands if s <= best_score + 160]
        if prefer_smaller_bay:
            # Among aligned dims, prefer bay half (3000) over overall (6000)
            return min(pool)
        return pool[0]

    a = best_leg(horiz_cands, prefer_smaller_bay=True)
    b = best_leg(vert_cands, prefer_smaller_bay=True)

    # Ensure a and b are distinct legs
    if a is not None and b is not None and abs(a - b) < 0.5:
        # Same value twice — try next distinct from the other band
        other_h = [v for _, v in sorted(horiz_cands) if abs(v - a) > 0.5]
        other_v = [v for _, v in sorted(vert_cands) if abs(v - b) > 0.5]
        if other_v:
            b = other_v[0]
        elif other_h:
            a = other_h[0]

    if a is None or b is None:
        # Fallback: two nearest distinct dims
        scored = sorted(
            (
                math.hypot(d["cx"] - mark["cx"], d["cy"] - mark["cy"]),
                float(d["text"]),
            )
            for d in dims
            if 900 <= float(d["text"]) <= 30000
        )
        unique: list[float] = []
        for _, val in scored:
            if all(abs(val - u) > 0.5 for u in unique):
                unique.append(val)
            if len(unique) >= 2:
                break
        if a is None and unique:
            a = unique[0]
        if b is None and len(unique) >= 2:
            # Prefer a value different from a; among leftovers prefer smaller (bay)
            rest = [v for v in unique[1:] if abs(v - float(a)) > 0.5]
            b = min(rest) if rest else unique[1]

    if a is None or b is None:
        return None
    if a < 900 or b < 900:
        return None
    return float(a), float(b)


def brace_diagonal_size(
    mark: dict[str, Any],
    dims: list[dict[str, Any]],
) -> tuple[str, str]:
    """
    Compute bracing Size via √(a² + b²).

    Returns (size_string, length_note) e.g. ("3354", "√(3000² + 1500²)").
    """
    legs = pick_diagonal_legs(mark, dims)
    if not legs:
        return "", ""
    a, b = legs
    # Normalize order: larger first in note for stability, keep values as found
    hyp = diagonal_length_mm(a, b)
    note = f"√({int(a)}² + {int(b)}²)"
    return format_mm_length(hyp), note


def instance_size_for_beam(mark: dict[str, Any], dims: list[dict[str, Any]]) -> str:
    """
    Size for one beam mark instance from nearby dims.

    Prefers local bay dims (1000 / 2000) when present so marks like B6 can
    split into multiple size rows (short stubs vs vertical bay pieces).
    """
    near: list[tuple[float, str]] = []
    for d in dims:
        dist = math.hypot(d["cx"] - mark["cx"], d["cy"] - mark["cy"])
        if dist < 420:
            near.append((dist, d["text"]))
    near.sort()
    vals = [v for _, v in near[:8]]
    # Explicit short / bay heights first (common multi-size marks)
    for preferred in ("1000", "2000", "1500", "3000"):
        if preferred in vals[:5]:
            return preferred
    # OCR near-misses for 1000
    if any(v in ("1050", "950") for v in vals[:4]):
        return "1000"
    # Fall back to band-aligned association (may prefer overall span)
    aligned = associate_dim_to_mark(mark, dims)
    if aligned:
        return aligned
    return vals[0] if vals else ""


def infer_size_quantities_from_plan_image(
    image: Image.Image,
) -> dict[str, collections.Counter]:
    """
    Per-instance size counts: { 'B6': Counter({'1000': 4, '2000': 26}), ... }.

    Use this when the same mark appears at more than one length on the plan.
    """
    marks, dims = ocr_tokens_with_boxes(image)
    by_mark: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    brace_notes: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)

    for m in marks:
        name = m["text"]
        if BRACE_RE.fullmatch(name):
            size, note = brace_diagonal_size(m, dims)
            if size:
                key = f"{size}|{note}" if note else size
                by_mark[name][key] += 1
                if note:
                    brace_notes[name][note] += 1
            continue
        if BEAM_RE.fullmatch(name) or PB_RE.fullmatch(name):
            size = instance_size_for_beam(m, dims)
            if size:
                by_mark[name][size] += 1

    return dict(by_mark)


def scale_size_counts_to_total(
    size_counts: collections.Counter,
    digital_total: int,
) -> dict[str, int]:
    """Proportionally scale OCR instance counts to the digital mark total."""
    if digital_total <= 0:
        return {}
    if not size_counts:
        return {"": digital_total}
    ocr_total = sum(size_counts.values())
    items = list(size_counts.items())
    allocated: dict[str, int] = {}
    remaining = digital_total
    for i, (size, n) in enumerate(items):
        if i == len(items) - 1:
            allocated[size] = max(0, remaining)
            remaining = 0
        else:
            q = max(1, round(digital_total * n / ocr_total))
            q = min(q, remaining - (len(items) - i - 1))  # leave ≥1 for each remaining
            q = max(1, q) if remaining > (len(items) - i - 1) else max(0, remaining)
            allocated[size] = q
            remaining -= q
    return allocated


def infer_sizes_from_plan_image(
    image: Image.Image,
) -> dict[str, str]:
    """
    Return {member_name: size} by OCR'ing plan dimensions near marks.

    - Straight beams/PB: size = nearby span dimension (mode / largest frequent)
    - Bracing (BR*): size = √(a² + b²) from nearby horizontal + vertical dims
      stored as "3354|√(3000² + 1500²)" when a formula note is available

    For multi-size marks (B6 @ 1000 and B6 @ 2000), prefer
    ``infer_size_quantities_from_plan_image`` instead.
    """
    qty_map = infer_size_quantities_from_plan_image(image)
    inferred: dict[str, str] = {}
    for mark, ctr in qty_map.items():
        if not ctr:
            continue
        if BRACE_RE.fullmatch(mark):
            # Most common hypotenuse|note
            inferred[mark] = ctr.most_common(1)[0][0]
            continue
        # Single representative size: prefer most common; if tie, keep both via qty API
        ranked = ctr.most_common()
        mode_val, mode_n = ranked[0]
        best = mode_val
        for val, n in ranked:
            if n >= max(1, mode_n // 2) and int(float(val)) > int(float(best)):
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
        if name not in sizes:
            continue
        raw = sizes[name]
        size, note = raw, ""
        if "|" in raw:
            size, note = raw.split("|", 1)
        if not row.get(size_key):
            row[size_key] = size
        if note:
            row["Length Note"] = note
            if not row.get("Length"):
                row["Length"] = size
    return members
