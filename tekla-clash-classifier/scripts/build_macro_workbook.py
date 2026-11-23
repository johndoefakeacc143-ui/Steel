#!/usr/bin/env python3
"""Build the Tekla Clash Classifier workbook shell.

True .xlsm with compiled vbaProject.bin cannot be authored on Linux without Excel.
After importing TeklaClashClassifier.bas on Windows:
  - Auto_Open / SetupClassifierUI places a Form Control RUN button
  - Click RUN → pick 3 files → SaveAs a NEW results workbook
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

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
    run_fill = PatternFill("solid", fgColor="C00000")
    run_font = Font(bold=True, color="FFFFFF", size=16)

    ws["A1"] = "Tekla Clash Classifier"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws.merge_cells("A1:D1")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A2"] = (
        "Click RUN. Excel asks for 3 files one by one "
        "(Universal Report → Bolt Checking Report → ClashCheck.log), "
        "then saves results to a NEW workbook. This classifier file stays intact."
    )
    ws["A2"].font = hint_font
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 48

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

    # Visual RUN placeholder (real Form Control button is created by SetupClassifierUI / Auto_Open)
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

    ws["A11"] = "Setup note"
    ws["A11"].font = label_font
    ws["B11"] = (
        "After importing TeklaClashClassifier.bas and saving as .xlsm, "
        "run macro SetupClassifierUI once (or reopen the file) to place the clickable Form Control RUN button "
        "on top of this area. Then click RUN."
    )
    ws["B11"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[11].height = 40

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    # --- Instructions ---
    ins = wb.create_sheet("Instructions")
    lines = [
        "How to enable and run (Windows Excel)",
        "",
        "A. One-time setup",
        "  1. Open this workbook in Excel.",
        "  2. Save As → Excel Macro-Enabled Workbook (*.xlsm).",
        "  3. Alt+F11 → File → Import File… → TeklaClashClassifier.bas",
        "     (or copy the VBA_Code sheet into a new Module).",
        "  4. Save, close VBA editor.",
        "  5. Alt+F8 → SetupClassifierUI → Run  (places the RUN Form Control button).",
        "     Or simply reopen the .xlsm (Auto_Open calls SetupClassifierUI).",
        "",
        "B. Allow macros",
        "  - Trust Center → Macro Settings → Disable all macros with notification",
        "  - Click Enable Content when prompted",
        "",
        "C. Run (every time)",
        "  1. Click the RUN button on the Control sheet.",
        "  2. Select Universal Report (.xlsx/.xls).",
        "  3. Select Bolt Checking Report (.xlsx/.xls/.csv).",
        "  4. Select ClashCheck.log (.log/.txt).",
        "  5. Choose where to Save the NEW results workbook.",
        "  6. Results open as a separate file with 4 tabs; this classifier .xlsm is unchanged.",
        "",
        "D. Output tabs",
        "  - Part_to_Part_Clash",
        "  - Part_to_Bolt_Clash",
        "  - Bolt_to_Bolt_Clash",
        "  - Duplicate members  (ClashCheck.log line contains 'inside member')",
        "",
        "E. Python alternative (optional)",
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
    vba_sheet["A1"] = (
        "Import TeklaClashClassifier.bas (preferred), or paste this module into VBA. "
        "Entry point: RunClashClassification. Button setup: SetupClassifierUI."
    )
    vba_sheet["A1"].font = Font(bold=True)
    for i, line in enumerate(bas_text.splitlines(), 2):
        vba_sheet.cell(i, 1, line)
    vba_sheet.column_dimensions["A"].width = 120

    # No result category tabs in the classifier workbook — results go to a new file.

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)


def try_write_xlsm_shell(xlsx_path: Path, xlsm_path: Path) -> str:
    """Place a .xlsm filename next to the .xlsx shell (no compiled VBA on Linux)."""
    shutil.copy2(xlsx_path, xlsm_path)
    return (
        "Wrote .xlsm filename copy of the xlsx shell (no compiled VBA binary). "
        "Import TeklaClashClassifier.bas in Excel, run SetupClassifierUI, Save As .xlsm."
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", type=Path, default=ROOT / "dist")
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
