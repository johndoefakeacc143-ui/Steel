# Steel Structure OCR Extractor

OCR-based takeoff tool for **steel structure drawings**. Upload a PDF, scanned PDF, or image; the app traces beams, columns, and base plates and exports an Excel workbook with three sheets.

## Excel output

| Sheet | Columns |
|-------|---------|
| **Beam** | Member Name, Quantity, Size (+ Length, Material, Page) |
| **Columns** | Member Name, Quantity, Size (+ Height, Material, Page) |
| **Base Plate** | Member Name, Quantity, Size (+ Thickness, anchors, Page) |

## What it accepts

- Digital PDFs (embedded text)
- Scanned / image-only PDFs
- PDFs that mix text and images
- Standalone images: PNG, JPG, TIFF, BMP, WEBP

## How it works

1. **Read** — digital text via pdfplumber; scanned pages rendered and OCR’d with Tesseract (OpenCV preprocessing)
2. **Trace** — regex extractors find member marks (`B1`, `C2`, `BP1`), section sizes (`W18x35`, `ISMB300`, …), quantities, and plate sizes
3. **Aggregate** — duplicate marks merged; quantities summed
4. **Export** — styled `.xlsx` with sheets **Beam**, **Columns**, **Base Plate**

## Prerequisites

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [Poppler](https://poppler.freedesktop.org/) (`pdftoppm`) for PDF → image

```bash
# Ubuntu / Debian
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils

# macOS
brew install tesseract poppler
```

## Setup

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env               # optional
```

## Web app

```bash
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

Open http://127.0.0.1:8000

- Drop a drawing → **Extract to Excel** → preview tables → **Download Excel**
- Health: http://127.0.0.1:8000/api/health

### API

```bash
# JSON preview + base64 Excel
curl -F "file=@drawing.pdf" http://127.0.0.1:8000/api/extract

# Stream .xlsx directly
curl -F "file=@drawing.pdf" -o takeoff.xlsx http://127.0.0.1:8000/api/download
```

## CLI

```bash
python cli.py drawing.pdf
python cli.py scan.png -o takeoff.xlsx --json
python cli.py plan.pdf --dpi 300
```

## Project layout

```
app.py                 # FastAPI entry
cli.py                 # Command-line takeoff
steel_ocr/
  reader.py            # PDF / image → text (digital + OCR)
  preprocess.py        # OpenCV cleanup for scans
  patterns.py          # Mark / section / plate regexes
  extractors.py        # Beam / column / base-plate tracing
  excel_export.py      # Aggregate + 3-sheet workbook
  pipeline.py          # End-to-end process_drawing()
static/index.html      # Upload UI
tests/                 # Extraction + Excel tests
```

## Tests

```bash
pytest -q
```

## Tips for better OCR

- Prefer 200–300 DPI scans; set `OCR_DPI=300` in `.env` for dense title blocks
- Member schedules with clear `B1 / C1 / BP1` marks extract most reliably
- Edit `steel_ocr/patterns.py` if your office uses non-standard mark prefixes

## License

Internal engineering tooling — adapt patterns to your drawing standards as needed.
