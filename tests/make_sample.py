"""Create a sample digital PDF schedule for end-to-end smoke tests."""

from __future__ import annotations

from pathlib import Path


def create_sample_pdf(path: Path) -> Path:
    """Generate a minimal text PDF with beam/column/base-plate schedule lines."""
    try:
        from pypdf import PdfWriter
        from pypdf.generic import (
            DecodedStreamObject,
            DictionaryObject,
            NameObject,
            NumberObject,
            ArrayObject,
        )
    except ImportError as exc:
        raise SystemExit("pypdf required") from exc

    # Use reportlab if available; otherwise write a simple pdfplumber-friendly PDF via pillow+pdf2
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        path.parent.mkdir(parents=True, exist_ok=True)
        c = canvas.Canvas(str(path), pagesize=A4)
        width, height = A4
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, height - 50, "STRUCTURAL FRAMING PLAN — SAMPLE TAKEOFF")
        c.setFont("Helvetica", 11)
        lines = [
            "BEAM SCHEDULE",
            "B1  W18x35  Length=6000mm  Qty 4  A992",
            "B2  W21x44  Length=7500mm  Qty 2  A992",
            "B3  ISMB300  Length=4500mm  Qty 6",
            "",
            "COLUMN SCHEDULE",
            "C1  W14x82  Height=4500mm  Qty 8",
            "C2  UC254x254x73  Height=6000mm  Qty 4",
            "",
            "BASE PLATE SCHEDULE",
            "BP1  500x500x25  4-M20  Qty 8",
            "BP2  PL 600x600x30  4 Nos M24  Qty 4",
        ]
        y = height - 90
        for line in lines:
            c.drawString(50, y, line)
            y -= 18
        c.showPage()
        c.save()
        return path
    except ImportError:
        pass

    # Fallback: write a plain-text "drawing" image as PNG for OCR smoke test
    from PIL import Image, ImageDraw, ImageFont

    img_path = path.with_suffix(".png")
    img = Image.new("RGB", (1200, 900), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.load_default()
    text = (
        "STRUCTURAL FRAMING PLAN\n\n"
        "B1 W18x35 Qty 4\n"
        "B2 W21x44 Qty 2\n"
        "C1 W14x82 Qty 8\n"
        "BP1 500x500x25 4-M20 Qty 8\n"
    )
    draw.multiline_text((40, 40), text, fill="black", font=font, spacing=8)
    img.save(img_path)
    return img_path


if __name__ == "__main__":
    out = Path(__file__).parent / "fixtures" / "sample_steel_schedule.pdf"
    created = create_sample_pdf(out)
    print("Created", created)
