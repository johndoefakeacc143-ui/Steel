"""High-capacity spatial OCR: tiled multi-pass Tesseract for large steel GA sheets."""

from __future__ import annotations

import collections
import logging
import math
import os
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

# Tunable via env — higher = slower but reads thin CAD text better
OCR_DPI = int(os.getenv("OCR_DPI", "320"))
OCR_UPSCALE = float(os.getenv("OCR_UPSCALE", "1.5"))
OCR_TILE_SIZE = int(os.getenv("OCR_TILE_SIZE", "2200"))  # px; 0 = no tiling
OCR_TILE_OVERLAP = int(os.getenv("OCR_TILE_OVERLAP", "280"))
OCR_MIN_CONF = float(os.getenv("OCR_MIN_CONF", "10"))
OCR_PSMS = [int(x) for x in os.getenv("OCR_PSMS", "11,6,4").split(",") if x.strip()]

OCR_FIXES = {
    "BRI": "BR1",
    "BRİ": "BR1",
    "BRA": "BR4",  # common misread of BR4
    "BRÁ": "BR4",
    "BRO": "BR4",
    "BA": "B4",
    "BG": "B6",
    "BS": "B5",
    "BE": "B8",
    "BQ": "B6",
    "BTL": "B7",
    "BT": "B7",
}


def _normalize_mark(raw: str) -> str:
    t = re.sub(r"[^A-Za-z0-9]", "", raw).upper()
    t = OCR_FIXES.get(t, t)
    # BR + digit with OCR garbage: BR4l, BR1|, etc.
    m = re.fullmatch(r"(BR)(\d)[A-Z]?", t)
    if m:
        return m.group(1) + m.group(2)
    if t.endswith("L") and BEAM_RE.fullmatch(t[:-1]):
        t = t[:-1]
    # B + digit + junk
    m2 = re.fullmatch(r"(B)(\d{1,2})[A-Z]?", t)
    if m2 and BEAM_RE.fullmatch(m2.group(1) + m2.group(2)):
        return m2.group(1) + m2.group(2)
    return t


def _dedupe(items: list[dict[str, Any]], tol: float = 22.0) -> list[dict[str, Any]]:
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


def _preprocess_variants(gray: np.ndarray, upscale: float) -> list[tuple[str, np.ndarray, float]]:
    """Several CAD-friendly binarisations / scales for Tesseract."""
    variants: list[tuple[str, np.ndarray, float]] = [("gray", gray, 1.0)]

    up = cv2.resize(gray, None, fx=upscale, fy=upscale, interpolation=cv2.INTER_CUBIC)
    den = cv2.fastNlMeansDenoising(up, h=6)

    adaptive = cv2.adaptiveThreshold(
        den, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 9
    )
    variants.append(("adaptive", adaptive, upscale))

    # Otsu — good for high-contrast plotted dims
    blur = cv2.GaussianBlur(den, (3, 3), 0)
    _, otsu = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(("otsu", otsu, upscale))

    # Slight morphological open to separate touching characters on dense GA
    kernel = np.ones((2, 2), np.uint8)
    opened = cv2.morphologyEx(adaptive, cv2.MORPH_OPEN, kernel, iterations=1)
    variants.append(("open", opened, upscale))

    # Inverted (white-on-dark title blocks / hatches)
    mean_val = float(np.mean(den))
    if mean_val < 140:
        variants.append(("inv", cv2.bitwise_not(adaptive), upscale))

    return variants


def _iter_tiles(
    width: int,
    height: int,
    tile: int,
    overlap: int,
) -> list[tuple[int, int, int, int]]:
    """Return (x0, y0, x1, y1) tiles covering the image with overlap."""
    if tile <= 0 or (width <= tile and height <= tile):
        return [(0, 0, width, height)]

    step = max(tile - overlap, tile // 2)
    boxes: list[tuple[int, int, int, int]] = []
    y = 0
    while y < height:
        x = 0
        y1 = min(y + tile, height)
        while x < width:
            x1 = min(x + tile, width)
            boxes.append((x, y, x1, y1))
            if x1 >= width:
                break
            x += step
        if y1 >= height:
            break
        y += step
    return boxes


def _ocr_array(
    arr: np.ndarray,
    *,
    scale: float,
    origin_x: float,
    origin_y: float,
    psms: list[int],
    min_conf: float,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    marks: list[dict[str, Any]] = []
    dims: list[dict[str, Any]] = []

    for psm in psms:
        config = f"--oem 3 --psm {psm}"
        try:
            df = pytesseract.image_to_data(
                arr, config=config, output_type=pytesseract.Output.DATAFRAME
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("OCR psm=%s failed: %s", psm, exc)
            continue
        if df is None or df.empty:
            continue
        df = df.dropna(subset=["text"])
        df["text"] = df["text"].astype(str).str.strip()
        df = df[(df["text"] != "") & (df["conf"].astype(float) >= min_conf)]

        for r in df.itertuples():
            raw = str(r.text).strip()
            t = _normalize_mark(raw)
            cx = origin_x + (float(r.left) + float(r.width) / 2) / scale
            cy = origin_y + (float(r.top) + float(r.height) / 2) / scale
            item = {"text": t, "cx": cx, "cy": cy, "conf": float(r.conf)}

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
                    dims.append({"text": str(val), "cx": cx, "cy": cy, "conf": float(r.conf)})

            # Also accept "1500mm" / "2000MM"
            m2 = re.fullmatch(r"(\d{3,5})\s*mm", raw, re.IGNORECASE)
            if m2:
                val = int(m2.group(1))
                if 900 <= val <= 30000:
                    dims.append({"text": str(val), "cx": cx, "cy": cy, "conf": float(r.conf)})

    return marks, dims


def ocr_tokens_with_boxes(
    image: Image.Image,
    *,
    upscale: float | None = None,
    tile_size: int | None = None,
    high_capacity: bool = True,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    High-capacity OCR of a plan image → (marks, dimensions) with pixel centers.

    When high_capacity=True (default):
      - multiple preprocess variants
      - multiple Tesseract page-segmentation modes
      - overlapping tiles on large sheets so small marks (BR4, 1500) are not lost
    """
    upscale = OCR_UPSCALE if upscale is None else upscale
    tile_size = OCR_TILE_SIZE if tile_size is None else tile_size
    psms = OCR_PSMS if high_capacity else [11]
    min_conf = OCR_MIN_CONF

    bgr = cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape[:2]

    tiles = _iter_tiles(w, h, tile_size if high_capacity else 0, OCR_TILE_OVERLAP)
    logger.info(
        "OCR capacity: size=%sx%s tiles=%s upscale=%.2f psms=%s",
        w,
        h,
        len(tiles),
        upscale,
        psms,
    )

    all_marks: list[dict[str, Any]] = []
    all_dims: list[dict[str, Any]] = []

    for x0, y0, x1, y1 in tiles:
        crop = gray[y0:y1, x0:x1]
        variants = _preprocess_variants(crop, upscale if high_capacity else 1.25)
        # Limit variants on many tiles to keep runtime reasonable
        if high_capacity and len(tiles) > 4:
            variants = variants[:3]  # gray + adaptive + otsu

        for _name, arr, scale in variants:
            marks, dims = _ocr_array(
                arr,
                scale=scale,
                origin_x=float(x0),
                origin_y=float(y0),
                psms=psms,
                min_conf=min_conf,
            )
            all_marks.extend(marks)
            all_dims.extend(dims)

    marks = _dedupe(all_marks, 28)
    dims = _dedupe(all_dims, 18)
    logger.info(
        "OCR result: %s marks (%s braces), %s dims",
        len(marks),
        sum(1 for m in marks if BRACE_RE.fullmatch(m["text"])),
        len(dims),
    )
    return marks, dims


def diagonal_length_mm(leg_a_mm: float, leg_b_mm: float) -> float:
    """Length of a diagonally placed brace: L = √(a² + b²)."""
    return math.hypot(float(leg_a_mm), float(leg_b_mm))


def format_mm_length(mm: float) -> str:
    """Format plan length as whole millimetres (engineering takeoff)."""
    return str(int(round(mm)))


def _classify_nearby_dims(
    mark: dict[str, Any], dims: list[dict[str, Any]]
) -> tuple[list[tuple[float, int, str]], list[tuple[float, int, str]], list[tuple[float, int, str]]]:
    horiz: list[tuple[float, int, str]] = []
    vert: list[tuple[float, int, str]] = []
    near: list[tuple[float, int, str]] = []

    for d in dims:
        dx = abs(d["cx"] - mark["cx"])
        dy = abs(d["cy"] - mark["cy"])
        dxy = math.hypot(dx, dy)
        val = int(d["text"])
        if dy < 120 and dx < 550:
            horiz.append((dy * 2.0 + dx * 0.15, -val, d["text"]))
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
        pool.sort(key=lambda g: (g[1], g[0]))
    return pool[0][2]


def associate_dim_to_mark(mark: dict[str, Any], dims: list[dict[str, Any]]) -> str:
    """Best nearby plan dimension for a straight member (prefer larger aligned span)."""
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
    For a diagonal brace (BR1/BR4…), pick horizontal run (a) and vertical rise (b).

    Returns (a_mm, b_mm) for L = √(a² + b²).
    """
    horiz_cands: list[tuple[float, float]] = []
    vert_cands: list[tuple[float, float]] = []

    for d in dims:
        dx = abs(d["cx"] - mark["cx"])
        dy = abs(d["cy"] - mark["cy"])
        val = float(d["text"])
        if val < 900:
            continue
        is_horiz_band = dy < 150 and dy >= 8 and dx < 650
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
            return min(pool)
        return pool[0]

    a = best_leg(horiz_cands, prefer_smaller_bay=True)
    b = best_leg(vert_cands, prefer_smaller_bay=True)

    if a is not None and b is not None and abs(a - b) < 0.5:
        other_h = [v for _, v in sorted(horiz_cands) if abs(v - a) > 0.5]
        other_v = [v for _, v in sorted(vert_cands) if abs(v - b) > 0.5]
        if other_v:
            b = other_v[0]
        elif other_h:
            a = other_h[0]

    if a is None or b is None:
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
            rest = [v for v in unique[1:] if abs(v - float(a)) > 0.5]
            b = min(rest) if rest else unique[1]

    if a is None or b is None or a < 900 or b < 900:
        return None
    return float(a), float(b)


def brace_diagonal_size(
    mark: dict[str, Any],
    dims: list[dict[str, Any]],
) -> tuple[str, str]:
    """Compute bracing Size via √(a² + b²). Returns (size, note)."""
    legs = pick_diagonal_legs(mark, dims)
    if not legs:
        return "", ""
    a, b = legs
    hyp = diagonal_length_mm(a, b)
    note = f"√({int(a)}² + {int(b)}²)"
    return format_mm_length(hyp), note


def instance_size_for_beam(mark: dict[str, Any], dims: list[dict[str, Any]]) -> str:
    """Size for one beam mark instance from nearby dims (supports multi-size marks)."""
    near: list[tuple[float, str]] = []
    for d in dims:
        dist = math.hypot(d["cx"] - mark["cx"], d["cy"] - mark["cy"])
        if dist < 420:
            near.append((dist, d["text"]))
    near.sort()
    vals = [v for _, v in near[:8]]
    for preferred in ("1000", "2000", "1500", "3000", "6000"):
        if preferred in vals[:5]:
            return preferred
    if any(v in ("1050", "950") for v in vals[:4]):
        return "1000"
    aligned = associate_dim_to_mark(mark, dims)
    if aligned:
        return aligned
    return vals[0] if vals else ""


def infer_size_quantities_from_plan_image(
    image: Image.Image,
    *,
    high_capacity: bool = True,
) -> dict[str, collections.Counter]:
    """Per-instance size counts from high-capacity OCR."""
    marks, dims = ocr_tokens_with_boxes(image, high_capacity=high_capacity)
    by_mark: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)

    for m in marks:
        name = m["text"]
        if BRACE_RE.fullmatch(name):
            size, note = brace_diagonal_size(m, dims)
            if size:
                key = f"{size}|{note}" if note else size
                by_mark[name][key] += 1
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
    """Proportionally scale OCR instance counts to the authoritative mark total."""
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
            leave = len(items) - i - 1
            q = max(1, round(digital_total * n / ocr_total))
            q = min(q, max(0, remaining - leave))
            if remaining > leave:
                q = max(1, q)
            allocated[size] = q
            remaining -= q
    return allocated


def infer_sizes_from_plan_image(
    image: Image.Image,
    *,
    high_capacity: bool = True,
) -> dict[str, str]:
    """Single size per mark (mode). Prefer infer_size_quantities for multi-size marks."""
    qty_map = infer_size_quantities_from_plan_image(image, high_capacity=high_capacity)
    inferred: dict[str, str] = {}
    for mark, ctr in qty_map.items():
        if not ctr:
            continue
        if BRACE_RE.fullmatch(mark):
            inferred[mark] = ctr.most_common(1)[0][0]
            continue
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
