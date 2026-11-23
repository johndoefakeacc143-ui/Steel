#!/usr/bin/env python3
"""Build a valid Tekla Clash Classifier .xlsx shell (no fake .xlsm).

Compiled VBA cannot be authored on Linux. Deliver:
  - tekla-clash-classifier.xlsx  (opens in Excel)
  - TeklaClashClassifier.bas     (import, then Save As .xlsm on Windows)

After importing VBA, run SetupClassifierUI once to place the RUN Form Control
button, or Alt+F8 -> RunClashClassification.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
BAS_PATH = ROOT / "vba" / "TeklaClashClassifier.bas"


def build_workbook(out_xlsx: Path, bas_text: str) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Control"
    title_fill = PatternFill("solid", fgColor="1F4E79")
    title_font = Font(bold=True, color="FFFFFF", size=16, name="Calibri")
    label_font = Font(bold=True, size=11)
    hint_font = Font(italic=True, color="666666", size=10)
    run_fill = PatternFill("solid", fgColor="C00000")
    run_font = Font(bold=True, color="FFFFFF", size=14)

    ws["A1"] = "Tekla Clash Classifier"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws.merge_cells("A1:D1")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A2"] = (
        "Setup once: Save As .xlsm -> Alt+F11 Import TeklaClashClassifier.bas -> "
        "Alt+F8 SetupClassifierUI (creates RUN button). "
        "Then click RUN (or Alt+F8 RunClashClassification). "
        "Excel asks for 3 files one by one; results save to a NEW workbook."
    )
    ws["A2"].font = hint_font
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 56

    ws["A3"] = "Output tabs"
    ws["A3"].font = label_font
    ws["B3"] = (
        "Part_to_Part_Clash | Part_to_Bolt_Clash | Bolt_to_Bolt_Clash | Duplicate members"
    )

    ws["A5"] = "Last status"
    ws["A5"].font = label_font
    ws["A6"] = "Last counts"
    ws["A6"].font = label_font
    ws["B5"] = ""
    ws["B6"] = ""

    # Placeholder until SetupClassifierUI creates a real Form Control button
    ws["A8"] = "RUN"
    ws["A8"].font = run_font
    ws["A8"].fill = run_fill
    ws["A8"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A8:B9")
    ws.row_dimensions[8].height = 22
    ws.row_dimensions[9].height = 22
    thin = Border(
        left=Side(style="medium", color="800000"),
        right=Side(style="medium", color="800000"),
        top=Side(style="medium", color="800000"),
        bottom=Side(style="medium", color="800000"),
    )
    for row in ws["A8:B9"]:
        for cell in row:
            cell.border = thin
            cell.fill = run_fill
            cell.font = run_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["A11"] = "Important"
    ws["A11"].font = label_font
    ws["B11"] = (
        "Use this .xlsx file (not .xlsm from chat). "
        "After importing the .bas module, Save As Excel Macro-Enabled Workbook (*.xlsm). "
        "If Excel says macros blocked: file Properties -> Unblock. "
        "Red RUN cell is a label until SetupClassifierUI places the real button."
    )
    ws["B11"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[11].height = 52

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    ins = wb.create_sheet("Instructions")
    lines = [
        "How to enable and run (Windows Excel)",
        "",
        "IMPORTANT: Open the .xlsx file. Do not use a renamed .xlsm from chat.",
        "",
        "A. One-time setup",
        "  1. Open tekla-clash-classifier.xlsx in Excel.",
        "  2. File -> Save As -> Excel Macro-Enabled Workbook (*.xlsm).",
        "  3. Alt+F11 -> File -> Import File... -> TeklaClashClassifier.bas",
        "  4. Save. If prompted about macros: Enable Content.",
        "  5. Alt+F8 -> SetupClassifierUI -> Run  (places the clickable RUN button).",
        "",
        "B. If SECURITY RISK / macros blocked",
        "  Close Excel. Right-click the .xlsm -> Properties -> Unblock -> OK. Re-open.",
        "",
        "C. Every run",
        "  1. Click RUN (or Alt+F8 -> RunClashClassification).",
        "  2. Select Universal Report (.xlsx/.xls).",
        "  3. Select Bolt Checking Report (.xlsx/.xls/.csv).",
        "  4. Select ClashCheck.log (.log/.txt).",
        "  5. Save the NEW results workbook when prompted.",
        "",
        "D. Output headers",
        "  Part_to_Part_Clash: GUID1, GUID2, PROFILE1, PROFILE2, Remarks",
        "  Part_to_Bolt_Clash: GUID1, GUID2, PROFILE1, bolt1, Remarks",
        "  Bolt_to_Bolt_Clash: GUID1, GUID2, bolt1, bolt2, Remarks",
        "  Duplicate members: GUID1, GUID2, PROFILE1, PROFILE2, Remarks",
    ]
    for i, line in enumerate(lines, 1):
        ins.cell(i, 1, line)
        if i == 1:
            ins.cell(i, 1).font = Font(bold=True, size=14)
    ins.column_dimensions["A"].width = 110

    vba_sheet = wb.create_sheet("VBA_Code")
    vba_sheet["A1"] = (
        "Import TeklaClashClassifier.bas (preferred), or paste into a Module. "
        "Entry point: RunClashClassification."
    )
    vba_sheet["A1"].font = Font(bold=True)
    for i, line in enumerate(bas_text.splitlines(), 2):
        vba_sheet.cell(i, 1, line)
    vba_sheet.column_dimensions["A"].width = 120

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)

    # Sanity: must reopen with openpyxl (valid OOXML)
    wb2 = load_workbook(out_xlsx)
    assert "Control" in wb2.sheetnames
    wb2.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", type=Path, default=ROOT / "dist")
    args = ap.parse_args()
    bas_text = BAS_PATH.read_text(encoding="utf-8")
    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = out_dir / "tekla-clash-classifier.xlsx"
    bas_out = out_dir / "TeklaClashClassifier.bas"
    xlsm_path = out_dir / "tekla-clash-classifier.xlsm"

    build_workbook(xlsx_path, bas_text)
    shutil.copy2(BAS_PATH, bas_out)

    # Remove any previous fake .xlsm so users are not handed a broken file
    if xlsm_path.exists():
        xlsm_path.unlink()
        print("Removed broken/fake", xlsm_path)

    print("Wrote", xlsx_path)
    print("Wrote", bas_out)
    print("Deliver .xlsx only. User Save As .xlsm after importing .bas.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
