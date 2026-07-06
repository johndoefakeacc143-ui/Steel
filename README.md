# Steel Structure PDF BOM Reader

Extract a Bill of Materials (BOM) from steel structure PDF drawings and export it to Excel.

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
│   ├── pdf_reader.py       # PDF text/table extraction (pdfplumber)
│   ├── image_processor.py    # Image preprocessing (OpenCV)
│   ├── ocr.py                # OCR for scanned drawings (pytesseract)
│   ├── bom_parser.py         # Steel member parsing (pandas)
│   └── excel_exporter.py     # Excel output
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
