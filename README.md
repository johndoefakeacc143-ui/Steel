# SteelDraw AI Extractor

Professional full-stack app that reads steel structure engineering PDF drawings like a steel detailer/modeler and extracts **Beams**, **Columns**, and **Base Plates** into a multi-sheet Excel workbook.

## Features

- Upload large PDFs (up to **500 MB**), processed **page by page**
- Auto-detects **digital vs scanned** pages; scanned pages use OpenCV preprocessing + Tesseract OCR
- Regex + optional OpenAI (LangChain) extraction for marks, sections, lengths, elevations, materials, base plates
- Excel output with 4 sheets: `Beams`, `Columns`, `BasePlates`, `Summary`
- React UI: drag-and-drop upload → progress screen → results preview + download

## Folder structure

```
/backend
  main.py
  requirements.txt
/frontend
  src/App.jsx
  package.json
.env
README.md
```

## Prerequisites

- Python 3.10+
- Node.js 18+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on the system
- OpenAI API key (optional but recommended for better AI extraction)
- Optional: Ghostscript + `camelot-py` (table extraction; not required)

### Install Tesseract (examples)

```bash
# Ubuntu / Debian
sudo apt-get update && sudo apt-get install -y tesseract-ocr

# macOS
brew install tesseract

# Windows
# Install from https://github.com/UB-Mannheim/tesseract/wiki
# and ensure `tesseract` is on your PATH
```

## Setup

### 1. Environment variables

Copy the example env file and add your OpenAI key:

```bash
cp .env.example .env
```

Edit `.env`:

```env
OPENAI_API_KEY=sk-your-real-key
OPENAI_MODEL=gpt-4o-mini
```

> Without an API key, the app still runs using **regex extraction** only.

### 2. Backend

```bash
cd backend
python -m venv .venv

# Windows:
.venv\Scripts\activate

# macOS / Linux:
source .venv/bin/activate

python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

API docs: http://127.0.0.1:8000/docs

> **Windows note:** `camelot-py[cv]` is intentionally omitted from `requirements.txt`
> because its `pdftopng` dependency often fails to install. The app works without it
> (pdfplumber + OCR + AI). To add Camelot later: `pip install camelot-py==0.11.0`
> and install Ghostscript.

### 3. Frontend

```bash
cd frontend
npm install
npm run dev
```

Open http://127.0.0.1:5173

The Vite dev server proxies `/api` to the FastAPI backend.

## API

### `POST /api/upload`

Accepts a PDF (`multipart/form-data` field name: `file`) and returns an Excel download.

### `POST /api/extract`

Same extraction pipeline, but returns JSON for the UI preview plus a base64 Excel payload for download.

### `GET /api/health`

Health check and whether `OPENAI_API_KEY` is configured.

## Excel sheets

| Sheet | Contents |
|-------|----------|
| Beams | Mark, Section Size, Length, Material, Start EL, End EL, Page |
| Columns | Mark, Section Size, Height, Base Elevation, Top Elevation, Material, Page |
| BasePlates | Mark, Plate Size, Thickness, Anchor Bolt Dia, Anchor Bolt Qty, Top of Concrete EL, Page |
| Summary | Total counts, min/max elevation, total beam length, material summary |

## Editing extraction regex

All patterns live near the top of `backend/main.py` under the comment:

```text
# REGEX PATTERNS — edit these to match your drawing standards
```

Key variables:

- `SECTION_SIZE_RE` — W18x35, ISMB400, UB/UC, HSS, etc.
- `BEAM_MARK_RE` / `COLUMN_MARK_RE` / `BASEPLATE_MARK_RE`
- `LENGTH_RE`, `ELEVATION_RE`, `MATERIAL_RE`
- `PLATE_SIZE_RE`, `ANCHOR_BOLT_RE`

## UI flow

1. **Upload** — drag & drop PDF, shows file name + size
2. **Choose pages** — pick which pages are Plan (beams/bracing) and which are Elevation (columns)
3. **Loading** — progress bar: “AI is reading your drawing…”
4. **Results** — Plan metrics + Elevation metrics, table tabs + **Download Excel**

## Notes

- Large scanned drawings are slower because each page is rendered and OCR’d.
- Camelot table extraction is optional and skipped automatically if not installed.
- CORS is open for local development; tighten `allow_origins` before production.
