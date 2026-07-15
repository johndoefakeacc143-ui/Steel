"""Unit tests for steel member extraction and Excel export (no OCR required)."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from openpyxl import load_workbook

from steel_ocr.excel_export import aggregate_members, build_excel_bytes, members_preview
from steel_ocr.extractors import extract_base_plates, extract_beams, extract_columns
from steel_ocr.reader import DocumentText, PageText, detect_page_type
from steel_ocr.extractors import extract_document_members

FIXTURE_TEXT = """
STRUCTURAL FRAMING PLAN — LEVEL 2

BEAM SCHEDULE
Mark    Section     Length      Qty   Material
B1      W18x35      6000mm      4     A992
B2      W21x44      7500mm      2     A992
B3      ISMB300     4500mm      6     IS 2062
BM-4    UB457x191x67  12'-6\"    1     S355

COLUMN SCHEDULE
Mark    Section     Height      Qty
C1      W14x82      4500mm      8
C2      UC254x254x73  6000mm    4
COL-3   HEB300      5500mm      2

BASE PLATE SCHEDULE
Mark    Plate Size       Anchor      Qty
BP1     500x500x25       4-M20       8
BP2     PL 600x600x30    4 Nos M24   4
BASE PLATE BP3  450x450x20  (4) Ø22  Qty=2

ELEVATION A-A
Column C4 W12x65 Height=3750mm Qty 2
"""


def test_extract_beams_from_schedule():
    beams = extract_beams(FIXTURE_TEXT, page_num=1)
    names = {b["Member Name"] for b in beams}
    assert "B1" in names
    assert "B2" in names
    assert "B3" in names
    assert "BM-4" in names or "BM4" in names
    b1 = next(b for b in beams if b["Member Name"] == "B1")
    assert "W18X35" in b1["Size"].upper().replace(" ", "")
    assert b1["Quantity"] == 4


def test_extract_columns_from_schedule():
    cols = extract_columns(FIXTURE_TEXT, page_num=1)
    names = {c["Member Name"] for c in cols}
    assert "C1" in names
    assert "C2" in names
    c1 = next(c for c in cols if c["Member Name"] == "C1")
    assert "W14X82" in c1["Size"].upper().replace(" ", "")
    assert c1["Quantity"] == 8


def test_extract_base_plates():
    plates = extract_base_plates(FIXTURE_TEXT, page_num=1)
    names = {p["Member Name"] for p in plates}
    assert any(n.startswith("BP") for n in names)
    bp1 = next(p for p in plates if p["Member Name"] == "BP1")
    assert "500" in bp1["Size"]
    assert bp1["Quantity"] == 8


def test_aggregate_sums_quantity():
    rows = [
        {"Member Name": "B1", "Size": "W18X35", "Quantity": 2},
        {"Member Name": "B1", "Size": "W18X35", "Quantity": 3},
        {"Member Name": "B2", "Size": "W21X44", "Quantity": 1},
    ]
    agg = aggregate_members(rows)
    by_name = {r["Member Name"]: r["Quantity"] for r in agg}
    assert by_name["B1"] == 5
    assert by_name["B2"] == 1


def test_excel_has_three_required_sheets_and_columns():
    beams = extract_beams(FIXTURE_TEXT, 1)
    cols = extract_columns(FIXTURE_TEXT, 1)
    plates = extract_base_plates(FIXTURE_TEXT, 1)
    data = build_excel_bytes(beams, cols, plates)
    wb = load_workbook(io.BytesIO(data))
    assert wb.sheetnames == ["Beam", "Columns", "Base Plate"]
    for sheet in wb.sheetnames:
        headers = [c.value for c in wb[sheet][1]]
        assert headers[:3] == ["Member Name", "Quantity", "Size"]
        assert wb[sheet].max_row >= 2  # header + at least one data row


def test_members_preview_keys():
    beams = extract_beams(FIXTURE_TEXT, 1)
    cols = extract_columns(FIXTURE_TEXT, 1)
    plates = extract_base_plates(FIXTURE_TEXT, 1)
    preview = members_preview(beams, cols, plates)
    assert "Beam" in preview and "Columns" in preview and "Base Plate" in preview
    assert preview["counts"]["beams"] >= 1


def test_detect_page_type():
    assert detect_page_type("FLOOR FRAMING PLAN LEVEL 2") == "Plan"
    assert detect_page_type("BRACE FRAME ELEVATION") == "Elevation"


def test_document_pipeline_text_only():
    doc = DocumentText(
        filename="fixture.txt",
        pages=[PageText(page_number=1, text=FIXTURE_TEXT, source="digital", page_type="Plan")],
    )
    members = extract_document_members(doc)
    assert members["beams"]
    assert members["columns"]
    assert members["base_plates"]


def test_cli_export_file(tmp_path: Path):
    """Write a text file and process via extractors + excel (skip full OCR)."""
    from steel_ocr.excel_export import build_excel_bytes
    from steel_ocr.extractors import extract_document_members
    from steel_ocr.reader import DocumentText, PageText

    doc = DocumentText(
        filename="sample.txt",
        pages=[PageText(1, FIXTURE_TEXT, "digital")],
    )
    m = extract_document_members(doc)
    out = tmp_path / "out.xlsx"
    out.write_bytes(build_excel_bytes(m["beams"], m["columns"], m["base_plates"]))
    assert out.exists() and out.stat().st_size > 1000
