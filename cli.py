#!/usr/bin/env python3
"""
CLI for steel structure OCR takeoff.

Examples:
  python cli.py drawing.pdf
  python cli.py scan.png -o takeoff.xlsx
  python cli.py plan.pdf --dpi 300 --json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from steel_ocr.pipeline import process_drawing, process_drawing_to_file


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="OCR steel drawings (PDF/images) → Excel (Beam, Columns, Base Plate)",
    )
    parser.add_argument("input", type=Path, help="PDF or image path")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx path (default: <input>_steel_takeoff.xlsx)",
    )
    parser.add_argument("--dpi", type=int, default=None, help="OCR render DPI (PDF scans)")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print member preview JSON to stdout",
    )
    parser.add_argument(
        "--no-details",
        action="store_true",
        help="Only Member Name / Quantity / Size columns (no Length/Material/Page)",
    )
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"File not found: {args.input}", file=sys.stderr)
        return 1

    result = process_drawing(
        args.input,
        dpi=args.dpi,
        include_details=not args.no_details,
    )
    out = args.output or Path(f"{args.input.stem}_steel_takeoff.xlsx")
    out.write_bytes(result["excel_bytes"])

    preview = result["preview"]
    counts = preview["counts"]
    print(f"Wrote: {out}")
    print(
        f"Beams: {counts['beams']} types ({counts['beam_quantity_total']} qty) | "
        f"Columns: {counts['columns']} types ({counts['column_quantity_total']} qty) | "
        f"Base plates: {counts['base_plates']} types ({counts['base_plate_quantity_total']} qty)"
    )
    if args.json:
        print(json.dumps(preview, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
