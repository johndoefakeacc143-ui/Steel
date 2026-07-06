# Steel Structure Drawing BOM Reader

Extract a Bill of Materials (BOM) from steel structure drawings and export it to Excel.

## Supported inputs

Reads **any** of these from the `input/` folder (mixed types are fine):

- **Text PDF** — `pdfplumber` text/tables + vector-character mark recovery.
- **Scanned / image PDF** — page images are OCR'd automatically.
- **Image files** — `.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.bmp`, `.webp` (OCR).
- **CAD** — `.dxf` directly; `.dwg` via the ODA File Converter.

Two engines are available (`--engine`):

- `classic` (default) — pdfplumber + OpenCV/Tesseract OCR + geometry-based mark and
  dimension recovery. Best for vector/text drawings.
- `ai` — a vision LLM reads the drawing image holistically, including faint/rotated
  text and printed schedule/quantity tables. Best for scanned drawings, image-only
  files, and drawings whose BOM lives in graphic tables. Requires an API key
  (`ANU_AI_API_KEY` / `OPENAI_API_KEY`; optional `ANU_AI_MODEL`, `ANU_AI_BASE_URL`).

```bash
python app.py                 # all inputs, classic engine
python app.py --engine ai     # vision-LLM engine (needs API key)
python app.py --file scan.png # a single image / pdf / dxf
```

## Requirements

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (for image-based drawings)

```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr

# macOS
brew install tesseract
```

## Setup

```bash
pip install -r requirements.txt
```

## Project Structure

```
├── app.py              # Main entry point
├── input/              # Place PDF drawings here
├── output/             # Generated Excel BOM files
├── src/
│   ├── input_reader.py     # Unified reader: PDF / image / DXF·DWG
│   ├── pdf_reader.py       # PDF text/table extraction (pdfplumber)
│   ├── image_processor.py  # Image preprocessing (OpenCV)
│   ├── ocr.py              # OCR for scanned drawings (pytesseract)
│   ├── dimension_ocr.py    # Length recovery from dimension lines (OCR)
│   ├── mark_recovery.py    # Recover rotated/stacked marks from geometry
│   ├── ai_reader.py        # Vision-LLM engine (--engine ai)
│   ├── sections.py         # Section unit-weight catalogue
│   ├── bom_parser.py       # Steel member parsing (pandas)
│   └── excel_exporter.py   # Excel output
└── requirements.txt
```

## Usage

1. Place one or more steel structure PDF files in the `input/` folder.
2. Run the application:

```bash
python app.py
```

Process a single PDF:

```bash
python app.py --pdf drawing.pdf
```

Disable OCR fallback:

```bash
python app.py --no-ocr
```

Excel BOM files are written to `output/` as `<filename>_BOM.xlsx`.

## BOM Output Columns

| Column       | Description                          |
|--------------|--------------------------------------|
| item_no      | Sequential item number               |
| mark         | Member designation (e.g. W12X26)     |
| description  | Member description                   |
| quantity     | Count of members                     |
| length       | Length if detected                   |
| grade        | Steel grade (A36, A992, etc.)        |
| weight       | Weight if present in source tables   |
| source_page  | PDF page number                      |

## Supported Member Types

- Wide flange: `W12x26`
- HSS: `HSS4x4x1/4`
- Angles: `L4x4x1/4`
- Channels: `C12x20.7`
- Tees: `WT6x25`
- Plates: `PL1/2`

The parser extracts members from PDF tables and text. When a page has little embedded text, OpenCV preprocessing and Tesseract OCR are used automatically.

## Length extraction from dimension lines

General-arrangement drawings usually show member lengths as dimension annotations
rather than in a table, and those labels are drawn as graphics (often rotated) that
are not in the PDF text layer. When OCR is enabled (the default), each page image is
OCR'd — including a 90°-rotated pass to read vertical dimensions — and every beam /
bracing mark is assigned the nearest dimension value (the most common one per mark).
For example, in `ABC.pdf` this recovers `Length = 2000` for beams such as `B7`.

These lengths are OCR-based estimates from the drawing geometry; base plates (`BP*`)
and plan-bay references (`PB*`) are not assigned a length. Use `--no-ocr` to disable
this (lengths then stay blank unless present in a table or text).

### Authoritative length overrides

Proximity-based estimation can pick the wrong dimension for a member (for example a
horizontal beam whose label sits next to the perpendicular bay dimension). To set
exact, authoritative lengths, provide a JSON map of `mark -> length_mm`. By default
the app reads `input/member_lengths.json` (or pass `--lengths-file path.json`):

```json
{
  "B4": 6000,
  "B7": 2000
}
```

A mark that exists at **several lengths** (e.g. the same beam mark used both as a
short vertical and a long horizontal member) can be split into groups. Each group
becomes its own BOM row:

```json
{
  "B3": [
    { "length": 6000, "quantity": 2 },
    { "length": 2000, "quantity": 8 },
    { "length": 1500, "quantity": 4 }
  ]
}
```

Overrides take priority over every estimate. A scalar value sets one length for all
of that mark; a list replaces the auto-detected rows with the listed groups. Marks
not listed keep their OCR/estimated length.
