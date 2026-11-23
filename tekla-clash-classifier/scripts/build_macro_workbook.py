#!/usr/bin/env python3
"""Build the Tekla Clash Classifier workbook shell with a real Form Control RUN button.

True compiled VBA (vbaProject.bin) cannot be authored on Linux without Excel.
This builder:
  1. Creates the Control / Instructions / VBA_Code sheets
  2. Injects a legacy VML Form Control button labeled RUN on Control,
     with FmlaMacro=RunClashClassification
  3. Copies .bas for the user to import once in Excel

After import + Enable Content, clicking RUN starts the macro.
"""

from __future__ import annotations

import argparse
import shutil
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
BAS_PATH = ROOT / "vba" / "TeklaClashClassifier.bas"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_ODR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

VML_BUTTON = r"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype>
 <v:shape id="btnRunClashClassification" o:spid="_x0000_s1025" type="#_x0000_t201"
  style="position:absolute;margin-left:10pt;margin-top:10pt;width:110pt;height:36pt;z-index:1;mso-wrap-style:tight"
  fillcolor="#C00000" strokecolor="#800000" o:button="t" filled="t" stroked="t">
  <v:fill color2="#C00000" o:detectmouseclick="t"/>
  <v:stroke weight="1pt"/>
  <o:lock v:ext="edit" rotation="t" text="f"/>
  <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
   <div style="text-align:center;font-family:Calibri;font-size:16pt;font-weight:bold;color:white;">RUN</div>
  </v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>0, 8, 7, 8, 1, 100, 9, 8</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:AutoFill>False</x:AutoFill>
   <x:FmlaMacro>[0]!RunClashClassification</x:FmlaMacro>
   <x:TextHAlign>Center</x:TextHAlign>
   <x:TextVAlign>Center</x:TextVAlign>
  </x:ClientData>
 </v:shape>
</xml>
"""


def build_workbook(out_xlsx: Path, bas_text: str) -> None:
    wb = Workbook()

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
        "Click the RUN button below. Excel asks for 3 files one by one "
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

    # Leave rows 8–9 clear for the embedded Form Control button (VML).
    ws["A8"] = ""
    ws["B8"] = ""
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 24

    ws["A11"] = "First-time setup"
    ws["A11"].font = label_font
    ws["B11"] = (
        "1) Save As .xlsm  2) Alt+F11 → Import TeklaClashClassifier.bas  "
        "3) Enable Content  4) Click RUN. "
        "(If the button is missing, Alt+F8 → SetupClassifierUI.)"
    )
    ws["B11"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[11].height = 40

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    ins = wb.create_sheet("Instructions")
    lines = [
        "How to enable and run (Windows Excel)",
        "",
        "A. One-time setup",
        "  1. Open this workbook in Excel.",
        "  2. Save As → Excel Macro-Enabled Workbook (*.xlsm).",
        "  3. Alt+F11 → File → Import File… → TeklaClashClassifier.bas",
        "  4. Save. Click Enable Content if prompted.",
        "  5. The Control sheet already has a RUN Form Control button.",
        "     If you do not see it: Alt+F8 → SetupClassifierUI → Run.",
        "",
        "B. Every run",
        "  1. Click RUN on the Control sheet.",
        "  2. Select Universal Report (.xlsx/.xls).",
        "  3. Select Bolt Checking Report (.xlsx/.xls/.csv).",
        "  4. Select ClashCheck.log (.log/.txt).",
        "  5. Save the NEW results workbook when prompted.",
        "",
        "C. Output tabs (in the new workbook only)",
        "  - Part_to_Part_Clash",
        "  - Part_to_Bolt_Clash",
        "  - Bolt_to_Bolt_Clash",
        "  - Duplicate members  (log line contains 'inside member')",
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
    inject_run_form_button(out_xlsx)


def _register_namespaces() -> None:
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("r", NS_ODR)
    ET.register_namespace(
        "mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"
    )
    ET.register_namespace(
        "x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
    )


def inject_run_form_button(xlsx_path: Path) -> None:
    """Embed a legacy Form Control button on the Control worksheet via VML."""
    _register_namespaces()
    buf = BytesIO()
    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(
        buf, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        names = set(zin.namelist())
        # Find Control sheet path from workbook.xml
        wb_xml = zin.read("xl/workbook.xml")
        root = ET.fromstring(wb_xml)
        sheet_path = "xl/worksheets/sheet1.xml"
        for sheet in root.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            if sheet.attrib.get("name") == "Control":
                rid = sheet.attrib.get(f"{{{NS_ODR}}}id")
                rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
                for rel in rels:
                    if rel.attrib.get("Id") == rid:
                        target = rel.attrib.get("Target", "").lstrip("/")
                        if not target.startswith("xl/"):
                            target = "xl/" + target
                        sheet_path = target
                        break
                break

        sheet_rels_path = (
            "xl/worksheets/_rels/" + Path(sheet_path).name + ".rels"
        )
        vml_path = "xl/drawings/vmlDrawing1.vml"
        vml_rel_id = "rIdRunBtn"

        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = _ensure_vml_content_type(data, vml_path)
            elif info.filename == sheet_path:
                data = _ensure_legacy_drawing(data, vml_rel_id)
            elif info.filename == sheet_rels_path:
                data = _ensure_vml_rel(data, vml_rel_id, vml_path)
            zout.writestr(info, data)

        if sheet_rels_path not in names:
            zout.writestr(
                sheet_rels_path,
                _new_sheet_rels(vml_rel_id, vml_path).encode("utf-8"),
            )
        if vml_path not in names:
            zout.writestr(vml_path, VML_BUTTON.encode("utf-8"))

    xlsx_path.write_bytes(buf.getvalue())


def _ensure_vml_content_type(data: bytes, vml_path: str) -> bytes:
    text = data.decode("utf-8")
    part = "/" + vml_path
    if "vmlDrawing" not in text:
        # Default Override + ensure Default for vml if missing
        if 'Extension="vml"' not in text:
            text = text.replace(
                "<Default Extension=",
                '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>\n<Default Extension=',
                1,
            )
            if 'Extension="vml"' not in text:
                text = text.replace(
                    "</Types>",
                    '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>',
                )
        if part not in text:
            text = text.replace(
                "</Types>",
                f'<Override PartName="{part}" '
                f'ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                f"</Types>",
            )
    return text.encode("utf-8")


def _ensure_legacy_drawing(data: bytes, rel_id: str) -> bytes:
    text = data.decode("utf-8")
    if "legacyDrawing" in text:
        return data
    tag = f'<legacyDrawing r:id="{rel_id}"/>'
    # Prefer before closing worksheet; after pageMargins / drawing if present
    if "</worksheet>" in text:
        text = text.replace("</worksheet>", tag + "</worksheet>")
    return text.encode("utf-8")


def _ensure_vml_rel(data: bytes, rel_id: str, vml_path: str) -> bytes:
    text = data.decode("utf-8")
    if "vmlDrawing" in text or rel_id in text:
        return data
    target = "../drawings/" + Path(vml_path).name
    rel = (
        f'<Relationship Id="{rel_id}" '
        f'Type="{NS_ODR}/vmlDrawing" '
        f'Target="{target}"/>'
    )
    text = text.replace("</Relationships>", rel + "</Relationships>")
    return text.encode("utf-8")


def _new_sheet_rels(rel_id: str, vml_path: str) -> str:
    target = "../drawings/" + Path(vml_path).name
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        f'<Relationship Id="{rel_id}" Type="{NS_ODR}/vmlDrawing" Target="{target}"/>'
        "</Relationships>"
    )


def try_write_xlsm_shell(xlsx_path: Path, xlsm_path: Path) -> str:
    shutil.copy2(xlsx_path, xlsm_path)
    return (
        "Wrote .xlsm copy with embedded RUN Form Control (VML). "
        "Import TeklaClashClassifier.bas in Excel, Enable Content, click RUN."
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

    # Verify VML landed
    with zipfile.ZipFile(xlsx_path) as z:
        assert "xl/drawings/vmlDrawing1.vml" in z.namelist(), z.namelist()
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
        assert "legacyDrawing" in sheet, sheet[-300:]
        vml = z.read("xl/drawings/vmlDrawing1.vml").decode()
        assert "RunClashClassification" in vml
        assert "btnRunClashClassification" in vml

    print("Wrote", xlsx_path)
    print("Wrote", xlsm_path, "-", note)
    print("Wrote", bas_out)
    print("Verified embedded RUN Form Control (VML) on Control sheet.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
