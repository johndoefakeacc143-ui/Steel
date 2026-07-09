"""Tests for along-beam dimensions and diagonal √(a²+b²) bracing lengths."""

from __future__ import annotations

import math
import sys
from pathlib import Path

import pdfplumber

# Allow `python backend/test_spatial_lengths.py` from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent))

from main import (  # noqa: E402
    DEFAULT_PLAN_BEAM_LENGTHS_MM,
    DEFAULT_PLAN_BRACE_LEGS_MM,
    apply_known_plan_lengths,
    diagonal_length_mm,
    extract_spatial_plan_lengths,
    process_pdf,
)


def test_diagonal_formula():
    assert abs(diagonal_length_mm(6000, 6000) - 6000 * math.sqrt(2)) < 0.01
    assert abs(diagonal_length_mm(6000, 2000) - math.hypot(6000, 2000)) < 0.01
    print("OK diagonal_length_mm")


def test_spatial_on_sample_plan():
    pdf_path = Path(__file__).parent / "sample_plan_with_dims.pdf"
    assert pdf_path.exists(), f"missing {pdf_path}"

    with pdfplumber.open(str(pdf_path)) as pdf:
        spatial = extract_spatial_plan_lengths(pdf.pages[0], 1)

    beams = {b["Mark"]: b for b in spatial["beams"]}
    bracing = {b["Mark"]: b for b in spatial["bracing"]}

    assert "B3" in beams, beams
    assert "B8" in beams, beams
    assert "B7" in beams, beams
    assert "B4" in beams, beams
    assert "BR1" in bracing, bracing

    assert beams["B3"]["Length"].startswith("1500"), beams["B3"]
    assert beams["B8"]["Length"].startswith("6000"), beams["B8"]
    assert beams["B7"]["Length"].startswith("2000"), beams["B7"]
    assert beams["B4"]["Length"].startswith("6000"), beams["B4"]

    # BR1 should use √(a²+b²) — sample has 45° + 6000 bay
    br1_len = bracing["BR1"]["Length"]
    assert br1_len, bracing["BR1"]
    assert "√" in bracing["BR1"].get("Length Note", "") or bracing["BR1"].get(
        "Length Source"
    ) == "diagonal_formula"
    expected = diagonal_length_mm(6000, 6000)
    # parse mm from "8485.3 mm" or "8485 mm"
    num = float(br1_len.replace("mm", "").strip())
    assert abs(num - expected) < 1.0, (br1_len, expected)
    print("OK spatial sample plan", {k: beams[k]["Length"] for k in beams}, bracing["BR1"])


def test_known_length_fallback():
    beams = [
        {"Mark": "B3", "Length": "", "Page": 1},
        {"Mark": "B8", "Length": "", "Page": 1},
        {"Mark": "B7", "Length": "", "Page": 1},
        {"Mark": "B4", "Length": "", "Page": 1},
        {"Mark": "B2", "Length": "", "Page": 1},
    ]
    bracing = [{"Mark": "BR1", "Length": "", "Page": 1}]
    apply_known_plan_lengths(
        beams,
        bracing,
        known_beam_lengths_mm=DEFAULT_PLAN_BEAM_LENGTHS_MM,
        known_brace_legs_mm=DEFAULT_PLAN_BRACE_LEGS_MM,
    )
    by_mark = {b["Mark"]: b["Length"] for b in beams}
    assert by_mark["B3"] == "1500 mm"
    assert by_mark["B8"] == "6000 mm"
    assert by_mark["B7"] == "2000 mm"
    assert by_mark["B4"] == "6000 mm"
    assert by_mark["B2"] == ""  # no default
    br = bracing[0]
    assert abs(float(br["Length"].replace("mm", "").strip()) - 6000 * math.sqrt(2)) < 1
    assert "√(6000² + 6000²)" in br["Length Note"]
    print("OK known length fallback", by_mark, br)


def test_process_pdf_end_to_end():
    pdf_path = Path(__file__).parent / "sample_plan_with_dims.pdf"
    _excel, meta = process_pdf(str(pdf_path), plan_pages=[1], elevation_pages=[])
    beams = {b["Mark"]: b for b in meta["beams"]}
    bracing = {b["Mark"]: b for b in meta["bracing"]}
    assert beams["B3"]["Length"].startswith("1500")
    assert beams["B8"]["Length"].startswith("6000")
    assert beams["B7"]["Length"].startswith("2000")
    assert beams["B4"]["Length"].startswith("6000")
    assert bracing["BR1"]["Length"]
    plan = meta["view_metrics"]["plan"]
    assert plan["beam_length_m"] > 0
    assert plan["bracing_length_m"] > 0
    print(
        "OK process_pdf",
        {k: beams[k]["Length"] for k in ("B3", "B8", "B7", "B4") if k in beams},
        bracing["BR1"]["Length"],
        plan,
    )


if __name__ == "__main__":
    test_diagonal_formula()
    test_spatial_on_sample_plan()
    test_known_length_fallback()
    test_process_pdf_end_to_end()
    print("\nAll spatial length tests passed.")
