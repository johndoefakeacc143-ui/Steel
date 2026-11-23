#!/usr/bin/env python3
"""Build the Tekla Clash Classifier workbook (xlsx shell + embedded VBA import instructions).

True .xlsm with a compiled vbaProject.bin cannot be reliably authored on Linux without Excel.
This builder creates a ready-to-use workbook shell. After importing the .bas module in Excel
on Windows, save as .xlsm. A copy is also written beside the .bas for distribution.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
BAS_PATH = ROOT / "vba" / "TeklaClashClassifier.bas"


def build_workbook(out_xlsx: Path, bas_text: str) -> None:
    wb = Workbook()

    # --- Control ---
    ws = wb.active
    ws.title = "Control"
    title_fill = PatternFill("solid", fgColor="1F4E79")
    title_font = Font(bold=True, color="FFFFFF", size=16, name="Calibri")
    label_font = Font(bold=True, size=11)
    hint_font = Font(italic=True, color="666666", size=10)

    ws["A1"] = "Tekla Clash Classifier"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws.merge_cells("A1:D1")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A2"] = (
        "1) Enable macros after you import VBA and save as .xlsm.  "
        "2) Browse/set the 3 input paths below.  "
        "3) Run macro: RunClashClassification (or Alt+F8)."
    )
    ws["A2"].font = hint_font
    ws.merge_cells("A2:D2")

    ws["A3"] = "Universal Report"
    ws["A4"] = "Bolt Report"
    ws["A5"] = "ClashCheck.log"
    for r in (3, 4, 5):
        ws.cell(r, 1).font = label_font
        ws.cell(r, 2).fill = PatternFill("solid", fgColor="FFF2CC")

    ws["A3"].font = label_font
    ws["B3"] = ""  # path
    ws["C3"] = "Browse macro: BrowseUniversal"
    ws["C4"] = "Browse macro: BrowseBolt"
    ws["C5"] = "Browse macro: BrowseClashLog"
    ws["A7"] = "Last status"
    ws["A8"] = "Last counts"
    ws["A7"].font = label_font
    ws["A8"].font = label_font

    ws["A10"] = "RUN"
    ws["A10"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A10"].fill = PatternFill("solid", fgColor="C00000")
    ws["B10"] = "After VBA import: Developer → Macros → RunClashClassification"
    ws["A11"] = "Output tabs created"
    ws["B11"] = "Part_to_Part_Clash | Part_to_Bolt_Clash | Bolt_to_Bolt_Clash | Duplicate members"
    ws["A12"] = "Column style"
    ws["B12"] = "GUID1 | GUID2 | (spacer) | attr1 (Profile or BOLT_DIA) | (spacer) | attr2 | Notes (blank for review)"

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 92
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 20

    # --- Instructions ---
    ins = wb.create_sheet("Instructions")
    lines = [
        "How to enable and run (Windows Excel)",
        "",
        "A. Import the VBA module (one-time setup)",
        "  1. Open this workbook in Excel.",
        "  2. Save As → Excel Macro-Enabled Workbook (*.xlsm).",
        "  3. Press Alt+F11 to open the VBA editor.",
        "  4. File → Import File… → select TeklaClashClassifier.bas (same folder / docs/vba).",
        "     Or copy all text from the VBA_Code sheet into a new Module.",
        "  5. Close the VBA editor and save the .xlsm.",
        "",
        "B. Allow macros",
        "  - File → Options → Trust Center → Trust Center Settings → Macro Settings",
        "  - Choose 'Disable all macros with notification' (recommended) or enable for this file.",
        "  - When you reopen the workbook, click Enable Content if prompted.",
        "",
        "C. Run classification",
        "  1. On the Control sheet, fill B3/B4/B5 with paths, or Alt+F8 → BrowseUniversal / BrowseBolt / BrowseClashLog.",
        "  2. Alt+F8 → RunClashClassification → Run.",
        "  3. Review sheets: Part_to_Part_Clash, Part_to_Bolt_Clash, Bolt_to_Bolt_Clash, Duplicate members.",
        "  4. Column G (Notes) is left blank for manual review comments (same idea as sample final tabs).",
        "",
        "D. Inputs",
        "  - Universal Report: .xlsx (GUID column auto-detected; title rows skipped).",
        "  - Bolt report: .xlsx or .csv (GUID values are Trimmed).",
        "  - ClashCheck.log: Tekla clash log lines 'Clash -> Guid: … <-> Guid: …'.",
        "  - Lines whose description contains 'inside member' go to 'Duplicate members'.",
        "",
        "E. Python alternative (no VBA)",
        "  python tekla-clash-classifier/scripts/classify_clashes.py \\",
        "    --universal Universal_Report.xlsx \\",
        "    --bolt BOLT_CHECKING_REPORT.xlsx \\",
        "    --clash ClashCheck.log \\",
        "    --output classified_clashes.xlsx",
    ]
    for i, line in enumerate(lines, 1):
        ins.cell(i, 1, line)
        if i == 1:
            ins.cell(i, 1).font = Font(bold=True, size=14)
    ins.column_dimensions["A"].width = 110

    # --- VBA_Code ---
    vba_sheet = wb.create_sheet("VBA_Code")
    vba_sheet["A1"] = "Paste this module into VBA (or import TeklaClashClassifier.bas). Do not edit paths here."
    vba_sheet["A1"].font = Font(bold=True)
    for i, line in enumerate(bas_text.splitlines(), 2):
        vba_sheet.cell(i, 1, line)
    vba_sheet.column_dimensions["A"].width = 120

    # Placeholder output sheets (cleared/rewritten by macro)
    for name in (
        "Part_to_Part_Clash",
        "Part_to_Bolt_Clash",
        "Bolt_to_Bolt_Clash",
        "Duplicate members",
    ):
        s = wb.create_sheet(name)
        s["A1"] = "GUID1"
        s["B1"] = "GUID2"
        s["A1"].font = Font(bold=True)
        s["B1"].font = Font(bold=True)
        s["A2"] = "(Run RunClashClassification to populate)"

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)


def try_write_xlsm_shell(xlsx_path: Path, xlsm_path: Path) -> str:
    """Place a .xlsm filename next to the .xlsx shell.

    A compiled vbaProject.bin cannot be authored on Linux without Excel, so this
    file is the same OpenXML workbook as the .xlsx. On Windows: open it (or the
    .xlsx), import TeklaClashClassifier.bas, then Save As *.xlsm to embed macros.
    """
    shutil.copy2(xlsx_path, xlsm_path)
    return (
        "Wrote .xlsm filename copy of the xlsx shell (no compiled VBA binary). "
        "Import TeklaClashClassifier.bas in Excel, then Save As .xlsm."
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "dist",
        help="Output directory for workbook + bas copy",
    )
    args = ap.parse_args()
    bas_text = BAS_PATH.read_text(encoding="utf-8")
    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = out_dir / "tekla-clash-classifier.xlsx"
    xlsm_path = out_dir / "tekla-clash-classifier.xlsm"
    bas_out = out_dir / "TeklaClashClassifier.bas"

    build_workbook(xlsx_path, bas_text)
    note = try_write_xlsm_shell(xlsx_path, xlsm_path)
    shutil.copy2(BAS_PATH, bas_out)
    print("Wrote", xlsx_path)
    print("Wrote", xlsm_path, "-", note)
    print("Wrote", bas_out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
