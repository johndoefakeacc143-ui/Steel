#!/usr/bin/env python3
"""Tekla clash classifier: Universal Report + Bolt report + ClashCheck.log -> category sheets."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

GUID_RE = re.compile(
    r"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
)
CLASH_RE = re.compile(
    r"Guid:\s*([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
    r"\s*<->\s*Guid:\s*"
    r"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})",
    re.IGNORECASE,
)

SHEET_PART_PART = "Part_to_Part_Clash"
SHEET_PART_BOLT = "Part_to_Bolt_Clash"
SHEET_BOLT_BOLT = "Bolt_to_Bolt_Clash"
SHEET_DUP = "Duplicate members"
SHEET_SUMMARY = "Summary"
SHEET_UNRESOLVED = "Unresolved"

OUTPUT_SHEETS = (
    SHEET_PART_PART,
    SHEET_PART_BOLT,
    SHEET_BOLT_BOLT,
    SHEET_DUP,
)


def normalize_guid(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.upper().startswith("GUID:"):
        text = text[5:]
    text = text.strip()
    m = GUID_RE.search(text)
    return m.group(1).lower() if m else text.lower()


def format_guid(guid: str) -> str:
    return f"GUID:{guid}" if guid else ""


def cell_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def looks_like_guid_header(value: Any) -> bool:
    text = cell_str(value).lower().replace(" ", "")
    return text in {"guid", "guid1", "guid2"} or text.startswith("guid")


def find_header_row_and_guid_col(
    rows: List[Tuple[Any, ...]], scan_limit: int = 40
) -> Tuple[int, int, Dict[str, int]]:
    """Return (header_row_index_0based, guid_col_index_0based, header_map)."""
    best: Optional[Tuple[int, int, Dict[str, int]]] = None
    for i, row in enumerate(rows[:scan_limit]):
        header_map: Dict[str, int] = {}
        guid_col = None
        for j, val in enumerate(row):
            if val is None:
                continue
            key = cell_str(val).lower()
            if not key:
                continue
            header_map[key] = j
            compact = key.replace(" ", "")
            if compact == "guid" or compact.startswith("guid"):
                # Prefer exact GUID over GUID1/GUID2 for source reports
                if compact == "guid":
                    guid_col = j
                elif guid_col is None and "guid" in compact:
                    guid_col = j
        if guid_col is not None and len(header_map) >= 2:
            # Prefer rows that also have Profile / BOLT_DIA style fields
            score = 0
            keys = set(header_map)
            if "profile" in keys:
                score += 5
            if any("bolt_dia" in k.replace(" ", "") for k in keys):
                score += 5
            if "name" in keys or "part pos" in keys:
                score += 2
            candidate = (score, -i, i, guid_col, header_map)
            if best is None or candidate[:2] > best[:2]:
                best = (candidate[0], candidate[1], candidate[2], candidate[3], candidate[4])
    if best is None:
        raise ValueError("Could not find a GUID header column in the report.")
    return best[2], best[3], best[4]


def header_index(header_map: Dict[str, int], *candidates: str) -> Optional[int]:
    normalized = {k.replace(" ", "").lower(): v for k, v in header_map.items()}
    for cand in candidates:
        key = cand.replace(" ", "").lower()
        if key in normalized:
            return normalized[key]
    # fuzzy contains
    for cand in candidates:
        key = cand.replace(" ", "").lower()
        for hk, hv in normalized.items():
            if key == hk or key in hk:
                return hv
    return None


@dataclass
class PartInfo:
    guid: str
    profile: str = ""
    name: str = ""
    raw: Dict[str, Any] = field(default_factory=dict)


@dataclass
class BoltInfo:
    guid: str
    diameter: Any = ""
    standard: str = ""
    raw: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ClashPair:
    guid1: str
    guid2: str
    line: str
    is_inside_member: bool
    source_order: int


def read_sheet_rows(path: Path) -> List[Tuple[Any, ...]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        rows: List[Tuple[Any, ...]] = []
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                rows.append(tuple(row))
        return rows

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        # Prefer a sheet whose name mentions universal/bolt/report; else active
        ws = wb.active
        for name in wb.sheetnames:
            low = name.lower()
            if "universal" in low or "bolt" in low or "report" in low:
                ws = wb[name]
                break
        return [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def load_parts(path: Path) -> Dict[str, PartInfo]:
    rows = read_sheet_rows(path)
    header_i, guid_col, header_map = find_header_row_and_guid_col(rows)
    profile_col = header_index(header_map, "profile")
    name_col = header_index(header_map, "name")
    parts: Dict[str, PartInfo] = {}
    for row in rows[header_i + 1 :]:
        if guid_col >= len(row):
            continue
        guid = normalize_guid(row[guid_col])
        if not guid or not GUID_RE.fullmatch(guid):
            # also accept if normalize found guid inside noisy cell
            if not guid or "-" not in guid:
                continue
        profile = cell_str(row[profile_col]) if profile_col is not None and profile_col < len(row) else ""
        name = cell_str(row[name_col]) if name_col is not None and name_col < len(row) else ""
        parts[guid] = PartInfo(guid=guid, profile=profile, name=name)
    return parts


def load_bolts(path: Path) -> Dict[str, BoltInfo]:
    rows = read_sheet_rows(path)
    header_i, guid_col, header_map = find_header_row_and_guid_col(rows)
    dia_col = header_index(header_map, "bolt_dia", "boltdia", "diameter", "dia")
    std_col = header_index(header_map, "bolt_std", "boltstd", "standard")
    bolts: Dict[str, BoltInfo] = {}
    for row in rows[header_i + 1 :]:
        if guid_col >= len(row):
            continue
        # CRITICAL: Trim spaces on bolt GUIDs
        guid = normalize_guid(row[guid_col])
        if not guid or "-" not in guid:
            continue
        dia: Any = ""
        if dia_col is not None and dia_col < len(row):
            dia = row[dia_col]
            if isinstance(dia, str):
                dia = dia.strip()
                if dia.isdigit():
                    dia = int(dia)
        std = cell_str(row[std_col]) if std_col is not None and std_col < len(row) else ""
        bolts[guid] = BoltInfo(guid=guid, diameter=dia, standard=std)
    return bolts


def parse_clash_log(path: Path) -> List[ClashPair]:
    clashes: List[ClashPair] = []
    with path.open("r", encoding="utf-8", errors="replace") as f:
        for order, line in enumerate(f):
            raw = line.rstrip("\n\r")
            m = CLASH_RE.search(raw)
            if not m:
                continue
            g1 = m.group(1).lower()
            g2 = m.group(2).lower()
            is_inside = "inside member" in raw.lower()
            clashes.append(
                ClashPair(
                    guid1=g1,
                    guid2=g2,
                    line=raw,
                    is_inside_member=is_inside,
                    source_order=order,
                )
            )
    return clashes


def kind_of(guid: str, parts: Dict[str, PartInfo], bolts: Dict[str, BoltInfo]) -> str:
    in_part = guid in parts
    in_bolt = guid in bolts
    if in_part and not in_bolt:
        return "part"
    if in_bolt and not in_part:
        return "bolt"
    if in_part and in_bolt:
        return "both"
    return "unknown"


def attr_for(
    guid: str, parts: Dict[str, PartInfo], bolts: Dict[str, BoltInfo]
) -> Any:
    if guid in parts:
        return parts[guid].profile
    if guid in bolts:
        return bolts[guid].diameter
    return ""


def classify(
    clashes: Iterable[ClashPair],
    parts: Dict[str, PartInfo],
    bolts: Dict[str, BoltInfo],
) -> Dict[str, List[Tuple[Any, ...]]]:
    buckets: Dict[str, List[Tuple[Any, ...]]] = {name: [] for name in OUTPUT_SHEETS}
    unresolved: List[Tuple[Any, ...]] = []

    for clash in clashes:
        k1 = kind_of(clash.guid1, parts, bolts)
        k2 = kind_of(clash.guid2, parts, bolts)

        # Duplicate / overlapping takes priority
        if clash.is_inside_member:
            buckets[SHEET_DUP].append(
                (
                    format_guid(clash.guid1),
                    format_guid(clash.guid2),
                    attr_for(clash.guid1, parts, bolts),
                    attr_for(clash.guid2, parts, bolts),
                    "",  # Remarks
                )
            )
            continue

        # Normalize part-bolt order: part first (matches sample PART TO BOLT)
        g1, g2, kk1, kk2 = clash.guid1, clash.guid2, k1, k2
        if kk1 == "bolt" and kk2 == "part":
            g1, g2 = g2, g1
            kk1, kk2 = kk2, kk1

        if kk1 == "part" and kk2 == "part":
            buckets[SHEET_PART_PART].append(
                (
                    format_guid(g1),
                    format_guid(g2),
                    attr_for(g1, parts, bolts),
                    attr_for(g2, parts, bolts),
                    "",
                )
            )
        elif kk1 == "part" and kk2 == "bolt":
            buckets[SHEET_PART_BOLT].append(
                (
                    format_guid(g1),
                    format_guid(g2),
                    attr_for(g1, parts, bolts),
                    attr_for(g2, parts, bolts),
                    "",
                )
            )
        elif kk1 == "bolt" and kk2 == "bolt":
            buckets[SHEET_BOLT_BOLT].append(
                (
                    format_guid(g1),
                    format_guid(g2),
                    attr_for(g1, parts, bolts),
                    attr_for(g2, parts, bolts),
                    "",
                )
            )
        else:
            unresolved.append(
                (
                    format_guid(clash.guid1),
                    format_guid(clash.guid2),
                    k1,
                    k2,
                    clash.line,
                )
            )

    buckets[SHEET_UNRESOLVED] = unresolved
    return buckets


CATEGORY_HEADERS = {
    SHEET_PART_PART: ["GUID1", "GUID2", "PROFILE1", "PROFILE2", "Remarks"],
    SHEET_PART_BOLT: ["GUID1", "GUID2", "PROFILE1", "bolt1", "Remarks"],
    SHEET_BOLT_BOLT: ["GUID1", "GUID2", "bolt1", "bolt2", "Remarks"],
    SHEET_DUP: ["GUID1", "GUID2", "PROFILE1", "PROFILE2", "Remarks"],
}


def style_header_row(ws, row: int, headers: List[str]) -> None:
    bold = Font(bold=True)
    for col, text in enumerate(headers, 1):
        cell = ws.cell(row, col, text)
        cell.font = bold


def write_category_sheet(wb: Workbook, name: str, rows: List[Tuple[Any, ...]]) -> None:
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    # Row 1: sheet title; Row 2: column headers; Row 3+: data
    title_cell = ws.cell(1, 1, name)
    title_cell.font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    headers = CATEGORY_HEADERS.get(name, ["GUID1", "GUID2", "PROFILE1", "PROFILE2", "Remarks"])
    style_header_row(ws, 2, headers)
    ws.freeze_panes = "A3"
    for r_i, row in enumerate(rows, 3):
        for c_i, val in enumerate(row, 1):
            ws.cell(r_i, c_i, val)
    widths = [42, 42, 22, 22, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_output_workbook(
    buckets: Dict[str, List[Tuple[Any, ...]]],
    out_path: Path,
    meta: Dict[str, Any],
) -> None:
    wb = Workbook()
    # Summary first
    ws = wb.active
    ws.title = SHEET_SUMMARY
    ws["A1"] = "Tekla Clash Classification Summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "Universal Report"
    ws["B3"] = meta.get("universal_path", "")
    ws["A4"] = "Bolt Report"
    ws["B4"] = meta.get("bolt_path", "")
    ws["A5"] = "ClashCheck.log"
    ws["B5"] = meta.get("clash_path", "")
    ws["A7"] = "Parts loaded"
    ws["B7"] = meta.get("parts_count", 0)
    ws["A8"] = "Bolts loaded"
    ws["B8"] = meta.get("bolts_count", 0)
    ws["A9"] = "Clash pairs parsed"
    ws["B9"] = meta.get("clash_count", 0)

    ws["A11"] = "Sheet"
    ws["B11"] = "Rows"
    ws["A11"].font = Font(bold=True)
    ws["B11"].font = Font(bold=True)
    for i, name in enumerate(OUTPUT_SHEETS, 12):
        ws.cell(i, 1, name)
        ws.cell(i, 2, len(buckets.get(name, [])))
    ws.cell(16, 1, SHEET_UNRESOLVED)
    ws.cell(16, 2, len(buckets.get(SHEET_UNRESOLVED, [])))
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80

    write_category_sheet(wb, SHEET_PART_PART, buckets[SHEET_PART_PART])
    write_category_sheet(wb, SHEET_PART_BOLT, buckets[SHEET_PART_BOLT])
    write_category_sheet(wb, SHEET_BOLT_BOLT, buckets[SHEET_BOLT_BOLT])
    write_category_sheet(wb, SHEET_DUP, buckets[SHEET_DUP])

    if buckets.get(SHEET_UNRESOLVED):
        if SHEET_UNRESOLVED in wb.sheetnames:
            del wb[SHEET_UNRESOLVED]
        ws_u = wb.create_sheet(SHEET_UNRESOLVED)
        ws_u.cell(1, 1, SHEET_UNRESOLVED).font = Font(bold=True, size=14)
        style_header_row(ws_u, 2, ["GUID1", "GUID2", "Kind1", "Kind2", "Log line"])
        for r_i, row in enumerate(buckets[SHEET_UNRESOLVED], 3):
            for c_i, val in enumerate(row, 1):
                ws_u.cell(r_i, c_i, val)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def run(universal: Path, bolt: Path, clash: Path, output: Path) -> Dict[str, Any]:
    parts = load_parts(universal)
    bolts = load_bolts(bolt)
    clashes = parse_clash_log(clash)
    buckets = classify(clashes, parts, bolts)
    meta = {
        "universal_path": str(universal),
        "bolt_path": str(bolt),
        "clash_path": str(clash),
        "parts_count": len(parts),
        "bolts_count": len(bolts),
        "clash_count": len(clashes),
        "counts": {k: len(buckets.get(k, [])) for k in list(OUTPUT_SHEETS) + [SHEET_UNRESOLVED]},
    }
    write_output_workbook(buckets, output, meta)
    return meta


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Classify Tekla clashes into category sheets.")
    p.add_argument("--universal", required=True, type=Path, help="Universal Report .xlsx")
    p.add_argument("--bolt", required=True, type=Path, help="Bolt checking report .xlsx/.csv")
    p.add_argument("--clash", required=True, type=Path, help="ClashCheck.log")
    p.add_argument("--output", required=True, type=Path, help="Output .xlsx path")
    args = p.parse_args(argv)
    meta = run(args.universal, args.bolt, args.clash, args.output)
    print("Wrote", args.output)
    for k, v in meta["counts"].items():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
