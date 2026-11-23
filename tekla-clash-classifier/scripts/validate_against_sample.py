#!/usr/bin/env python3
"""Validate classifier output against sample Universal_Report final/intermediate tabs."""

from __future__ import annotations

import re
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from classify_clashes import (  # noqa: E402
    SHEET_BOLT_BOLT,
    SHEET_PART_BOLT,
    SHEET_PART_PART,
    normalize_guid,
    run,
)

INPUTS = Path("/cursor/stores/self/internal/tekla-inputs")
OUT = Path("/tmp/tekla_clash_validation_out.xlsx")


def _find_guid_cols(row) -> tuple:
    """Return (guid1_idx, guid2_idx) for a data row, or (-1, -1)."""
    found = []
    for i, v in enumerate(row):
        g = normalize_guid(v)
        if g and "-" in g and g not in {"guid1", "guid2", "guid"}:
            # Prefer cells that look like UUID text
            s = str(v or "")
            if "guid" in s.lower() or len(g) >= 32:
                found.append(i)
        if len(found) >= 2:
            break
    if len(found) >= 2:
        return found[0], found[1]
    return -1, -1


def load_pairs(path: Path, sheet: str) -> dict:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    pairs = {}
    for row in ws.iter_rows(values_only=True):
        gi, gj = _find_guid_cols(row)
        if gi < 0:
            # legacy: GUID in col A/B
            a = normalize_guid(row[0] if row else None)
            b = normalize_guid(row[1] if row and len(row) > 1 else None)
            if not a or a in {"guid1", "guid"} or "-" not in a:
                continue
            gi, gj = 0, 1
        else:
            a = normalize_guid(row[gi])
            b = normalize_guid(row[gj])
            if not a or "-" not in a:
                continue
        key = frozenset([a, b])
        # New layout: attrs at F/H (5/7); mid layout attrs at 2/3; old spacer at 3/5
        if len(row) >= 9 and row[5] not in (None, "", " ") and str(row[5]).strip() not in {"PROFILE1", "bolt1"}:
            attr1, attr2 = row[5], row[7]
            remarks = row[8] if len(row) > 8 else None
        elif len(row) >= 5 and row[2] not in (None, "", " ") and "GUID" not in str(row[2]).upper():
            attr1, attr2 = row[2], row[3]
            remarks = row[4] if len(row) > 4 else None
        else:
            attr1 = row[3] if len(row) > 3 else None
            attr2 = row[5] if len(row) > 5 else None
            remarks = row[6] if len(row) > 6 else (row[4] if len(row) > 4 else None)
        pairs[key] = {
            "a": a,
            "b": b,
            "d": attr1,
            "f": attr2,
            "g": remarks,
            "a_fmt": row[gi],
            "b_fmt": row[gj],
        }
    wb.close()
    return pairs


def main() -> int:
    meta = run(
        INPUTS / "Universal_Report.xlsx",
        INPUTS / "BOLT_CHECKING_REPORT.xlsx",
        INPUTS / "ClashCheck.log",
        OUT,
    )
    print("Classifier counts:", meta["counts"])

    sample = INPUTS / "Universal_Report.xlsx"
    sample_ptb = load_pairs(sample, "PART TO BOLT")
    sample_ptp = load_pairs(sample, "PART TO PART")
    sample_final_ptb = load_pairs(sample, "final part to bolt check")
    sample_final_ptp = load_pairs(sample, "part to part final")

    our_ptb = load_pairs(OUT, SHEET_PART_BOLT)
    our_ptp = load_pairs(OUT, SHEET_PART_PART)

    print(f"Sample PART TO BOLT: {len(sample_ptb)} | Ours: {len(our_ptb)}")
    print(f"Sample PART TO PART: {len(sample_ptp)} | Ours: {len(our_ptp)}")
    print(f"Intersection PTB: {len(sample_ptb.keys() & our_ptb.keys())}")
    print(f"Intersection PTP: {len(sample_ptp.keys() & our_ptp.keys())}")
    print(f"Missing PTB (in sample not ours): {len(sample_ptb.keys() - our_ptb.keys())}")
    print(f"Extra PTB (ours not sample): {len(our_ptb.keys() - sample_ptb.keys())}")
    print(f"Missing PTP: {len(sample_ptp.keys() - our_ptp.keys())}")
    print(f"Extra PTP: {len(our_ptp.keys() - sample_ptp.keys())}")

    # Spot-check attribute values for intersecting pairs
    ptb_attr_ok = 0
    ptb_attr_bad = []
    for key in list(sample_ptb.keys() & our_ptb.keys())[:50]:
        s, o = sample_ptb[key], our_ptb[key]
        sd = str(s["d"]).strip() if s["d"] is not None else ""
        od = str(o["d"]).strip() if o["d"] is not None else ""
        sf = str(s["f"]).strip() if s["f"] is not None else ""
        of_ = str(o["f"]).strip() if o["f"] is not None else ""
        # Sample may have part on either side of frozenset; compare attrs as sets of stringified values
        if {sd, sf} == {od, of_} or (sd == od and sf == of_):
            ptb_attr_ok += 1
        else:
            # Also accept numeric dia equality
            try:
                if sd == od and float(sf) == float(of_):
                    ptb_attr_ok += 1
                    continue
            except (TypeError, ValueError):
                pass
            ptb_attr_bad.append((key, s, o))

    ptp_attr_ok = 0
    ptp_attr_bad = []
    for key in list(sample_ptp.keys() & our_ptp.keys())[:50]:
        s, o = sample_ptp[key], our_ptp[key]
        sd = str(s["d"]).strip() if s["d"] is not None else ""
        od = str(o["d"]).strip() if o["d"] is not None else ""
        sf = str(s["f"]).strip() if s["f"] is not None else ""
        of_ = str(o["f"]).strip() if o["f"] is not None else ""
        if {sd, sf} == {od, of_}:
            ptp_attr_ok += 1
        else:
            ptp_attr_bad.append((key, sd, sf, od, of_))

    print(f"PTB attr spot-check OK {ptb_attr_ok}/50 bad={len(ptb_attr_bad)}")
    print(f"PTP attr spot-check OK {ptp_attr_ok}/50 bad={len(ptp_attr_bad)}")
    if ptb_attr_bad[:3]:
        print("PTB bad sample", ptb_attr_bad[0])
    if ptp_attr_bad[:3]:
        print("PTP bad sample", ptp_attr_bad[0])

    # Final tabs should be subsets of our classified sets
    final_ptb_cover = len(sample_final_ptb.keys() & our_ptb.keys())
    final_ptp_cover = len(sample_final_ptp.keys() & our_ptp.keys())
    print(
        f"Final PTB covered {final_ptb_cover}/{len(sample_final_ptb)} | "
        f"Final PTP covered {final_ptp_cover}/{len(sample_final_ptp)}"
    )

    # Raw parser smoke tests
    from classify_clashes import load_bolts, load_parts

    raw_parts = load_parts(INPUTS / "sample_raw_Universal_Report.xlsx")
    raw_bolts = load_bolts(INPUTS / "sample_raw_BOLT_CHECKING_REPORT.CSV")
    print(f"Raw universal parts loaded: {len(raw_parts)} (GUID not col1)")
    print(f"Raw bolt CSV bolts loaded: {len(raw_bolts)} (Trim GUID)")
    # Check a bolt GUID had spaces in CSV
    with open(INPUTS / "sample_raw_BOLT_CHECKING_REPORT.CSV", encoding="utf-8") as f:
        second = f.readline()
        second = f.readline()
    raw_guid_field = second.split(",")[0]
    print(f"Raw bolt GUID field repr: {raw_guid_field!r}")
    print(f"Normalized in map: {normalize_guid(raw_guid_field)!r} in bolts={normalize_guid(raw_guid_field) in raw_bolts}")

    # Order check: part first in part-bolt + header titles
    wb = load_workbook(OUT, read_only=True, data_only=True)
    ws = wb[SHEET_PART_BOLT]
    headers = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
    print(f"Part_to_Bolt headers row2: {headers}")
    expect_pb = ["SL.No.", "-", "GUID1", "GUID2", "-", "PROFILE1", "-", "bolt1", "Remarks"]
    assert headers[:9] == expect_pb, headers
    title = ws.cell(1, 1).value
    print(f"Part_to_Bolt title row1: {title}")
    assert title == SHEET_PART_BOLT, title
    ws_bb = wb[SHEET_BOLT_BOLT]
    bb_headers = [c.value for c in next(ws_bb.iter_rows(min_row=2, max_row=2))]
    print(f"Bolt_to_Bolt headers row2: {bb_headers}")
    expect_bb = ["SL.No.", "-", "GUID1", "GUID2", "-", "bolt1", "-", "bolt2", "Remarks"]
    assert bb_headers[:9] == expect_bb, bb_headers

    part_first = 0
    checked = 0
    from classify_clashes import load_parts as lp, load_bolts as lb

    parts = lp(INPUTS / "Universal_Report.xlsx")
    bolts = lb(INPUTS / "BOLT_CHECKING_REPORT.xlsx")
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue  # title + header
        a = normalize_guid(row[2])
        b = normalize_guid(row[3])
        if a in parts and b in bolts:
            part_first += 1
        checked += 1
        if checked >= 100:
            break
    wb.close()
    print(f"Part-first in first 100 Part_to_Bolt rows: {part_first}/{checked}")

    # Sorted by PROFILE1/bolt1 then PROFILE2/bolt2 (whole rows) — cols F/H
    def is_sorted_by_attrs(sheet_name: str) -> bool:
        wbo = load_workbook(OUT, read_only=True, data_only=True)
        vals = []
        for i, row in enumerate(wbo[sheet_name].iter_rows(values_only=True)):
            if i < 2:
                continue
            if row[2] is None:
                continue
            vals.append(
                (
                    "" if row[5] is None else str(row[5]).strip().upper(),
                    "" if row[7] is None else str(row[7]).strip().upper(),
                )
            )
        wbo.close()
        return vals == sorted(vals)

    sort_pp = is_sorted_by_attrs(SHEET_PART_PART)
    sort_pb = is_sorted_by_attrs(SHEET_PART_BOLT)
    print(f"Sorted Part_to_Part by PROFILE1/PROFILE2: {sort_pp}")
    print(f"Sorted Part_to_Bolt by PROFILE1/bolt1: {sort_pb}")

    ok = (
        sort_pp
        and sort_pb
        and len(sample_ptb.keys() - our_ptb.keys()) == 0
        and len(sample_ptp.keys() - our_ptp.keys()) == 0
        and final_ptb_cover == len(sample_final_ptb)
        and final_ptp_cover == len(sample_final_ptp)
        and ptb_attr_ok >= 45
        and ptp_attr_ok >= 45
        and len(raw_parts) > 0
        and len(raw_bolts) > 0
    )
    print("VALIDATION", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
