#!/usr/bin/env python3
"""Generate a sample steel member-schedule PDF for testing the BOM reader.

Requires reportlab (a dev-only dependency, not in requirements.txt):
    pip install reportlab

Writes input/SAMPLE_STRUCTURE.pdf with a member-schedule table (and one free-text
member line) so the classic table/text extraction paths can be exercised.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUTPUT = Path(__file__).resolve().parent.parent / "input" / "SAMPLE_STRUCTURE.pdf"

DATA = [
    ["Member ID", "Type", "Section", "Length", "Qty", "Elevation (m)", "Grid", "Weight (kg)"],
    ["B1", "Beam", "ISMB 300", "6.0 m", "4", "12.5", "A-1", ""],
    ["C1", "Column", "ISMB 400", "4500 mm", "8", "0.0", "B-2", ""],
    ["BR1", "Bracing", "ISMC 150", "3.2 m", "6", "8.0", "C-3", ""],
    ["G1", "Girder", "ISMB 500", "9000 mm", "2", "16.0", "D-4", "782.1"],
    ["PL1", "Plate", "PL 200x10", "1500 mm", "10", "", "", ""],
    ["RHS1", "Bracing", "RHS 100x100x4", "2.5 m", "12", "5.0", "E-1", ""],
    ["B3", "Beam", "", "4.0 m", "2", "10.0", "F-2", ""],
]


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(OUTPUT), pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    story = [
        Paragraph("STEEL STRUCTURE - MEMBER SCHEDULE (PIPE RACK 400PB)", styles["Title"]),
        Spacer(1, 12),
        Paragraph("General notes: EL. datum +100.000 m. Grades to IS 2062.", styles["Normal"]),
        Spacer(1, 12),
    ]
    table = Table(DATA, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#305496")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#DCE6F1")]),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 18))
    story.append(
        Paragraph("Additional: B2 ISMB 200 Length 5.0 m Qty 3 Grid A-2 EL 12.5", styles["Normal"])
    )
    doc.build(story)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
