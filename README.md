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

**Every upload is scanned fresh.** Nothing is hardcoded from a previous drawing.

1. **Digital text** — pdfplumber counts marks on the text layer (when present)
2. **High-capacity OCR** — page rendered at high DPI, tiled, multi-pass Tesseract; reads marks and nearby dimensions (including graphics-only labels like some braces)
3. **Diagonal braces** — length \(L=\sqrt{a^2+b^2}\) from nearby horizontal run × vertical rise
4. **Optional vision** — if `OPENAI_API_KEY` or `GEMINI_API_KEY` is set, a vision model also reads the sheet like a human and fills OCR gaps
5. **Export** — `.xlsx` with **Beam**, **Columns**, **Base Plate** (Member Name, Quantity, Size). Same mark at different lengths → separate rows

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

- Prefer 200–300 DPI scans; set `OCR_DPI=320` in `.env` for dense title blocks
- Member schedules with clear `B1 / C1 / BP1` marks extract most reliably
- Edit `steel_ocr/patterns.py` if your office uses non-standard mark prefixes

## Free vision API key (Gemini)

OCR works with **no API key**. For human-like vision assist on each upload:

1. Open [Google AI Studio](https://aistudio.google.com/apikey) and create a **free** API key
2. Put it in `.env` (already created locally; not committed to git):

```env
GEMINI_API_KEY=your_key_here
GEMINI_VISION_MODEL=gemini-2.5-flash
```

3. Restart: `uvicorn app:app --host 0.0.0.0 --port 8000`
4. Check http://127.0.0.1:8000/api/health → `"vision_assist": true`

```bash
./setup_api_key.sh
```

A real key can only be created in your Google account — it cannot be generated inside this repo.

## License

Internal engineering tooling — adapt patterns to your drawing standards as needed.
