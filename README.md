# SteelDraw AI Extractor

Full-stack app that reads steel structure engineering PDF drawings like a detailer/modeler and extracts **Beams**, **Columns**, and **Base Plates** into an Excel workbook with four sheets: Beams, Columns, BasePlates, and Summary.

## Features

- Upload PDFs up to **500 MB**, processed **page by page** to limit memory use
- Detects **digital vs scanned** pages; scanned pages use OpenCV preprocessing + Tesseract OCR
- Optional **Camelot** table extraction for member schedules
- **Regex + OpenAI (LangChain)** extraction for marks, section sizes, lengths, elevations, materials
- **Page selection** when the drawing has more than 5 pages:
  - Plan pages → beams & bracings
  - Elevation pages → columns & base plates
- Drawings with **1–5 pages** scan automatically with no page prompt
- Excel download with engineer-style **Summary** (counts by size/length/weight)

## Folder structure

```
/backend
  main.py              # FastAPI app + extraction pipeline
  requirements.txt
/frontend
  index.html
  package.json
  vite.config.js
  tailwind.config.js
  src/App.jsx          # Upload → Loading → Results UI
  src/main.jsx
  src/index.css
.env.example
README.md
```

## Prerequisites

- Python 3.10+
- Node.js 18+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on the system
- Optional (for Camelot lattice tables): Ghostscript

### Install Tesseract (examples)

```bash
# Ubuntu / Debian
sudo apt-get update && sudo apt-get install -y tesseract-ocr ghostscript

# macOS
brew install tesseract ghostscript
```

## Setup

### 1. Environment variables

```bash
cp .env.example .env
# Edit the project-root .env (NOT only a random copy elsewhere):
#   OPENAI_API_KEY=sk-proj-your-real-key
# No quotes. One line. Then restart uvicorn.
```

Confirm AI is on: open [http://127.0.0.1:8000/api/health](http://127.0.0.1:8000/api/health) — `"openai_configured": true`.  
On startup the backend prints `[startup] OpenAI READY` or `OpenAI OFF`.

Regex/OCR extraction works without an API key. OpenAI improves mark association and the Summary narrative.

### 2. Backend

```bash
cd backend
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

Health check: [http://127.0.0.1:8000/api/health](http://127.0.0.1:8000/api/health)

### 3. Frontend

```bash
cd frontend
npm install
npm run dev
```

Open [http://127.0.0.1:5173](http://127.0.0.1:5173). Vite proxies `/api` to the backend.

## API

### `POST /api/upload`

Multipart form fields:

| Field | Description |
|-------|-------------|
| `file` | PDF (required on first upload) |
| `job_id` | Returned when page selection is needed |
| `plan_pages` | e.g. `1,2` or `1-3` (beams & bracings) |
| `elevation_pages` | e.g. `4-6` (columns & base plates) |
| `confirm` | `true` after the user picks pages |

**Response (≤5 pages or after confirm):** JSON with `preview` tables, `counts`, and `excel_base64` for download.

**Response (>5 pages, not confirmed):** `{ needs_page_selection: true, job_id, page_count, message }`

### `POST /api/download`

Same extraction, streams the `.xlsx` file directly (for API clients).

## Excel sheets

1. **Beams** — Mark, Section Size, Length, Material, Start EL, End EL  
2. **Columns** — Mark, Section Size, Height, Base/Top Elevation, Material  
3. **BasePlates** — Mark, Plate Size, Thickness, Anchor bolts, TOC EL, Weight  
4. **Summary** — Totals, min/max elevation, total beam length, material mix, and engineer counts such as “N columns of length L”, “N bracings of length L”, “N beams of length L”, “N base plates of weight W”

## Length association (plan drawings)

On plan views, member length is taken from the dimension written in the **same direction** as the member:

| Mark | Direction | Length |
|------|-----------|--------|
| B3   | vertical dim beside beam | 1500 |
| B8   | vertical dim (each instance) | 6000 |
| B7   | vertical dim | 2000 |
| B4   | horizontal dim | 6000 |

**Diagonal members** (BR1, BR4, or any beam/column marked diagonal/sloping) use the triangle diagonal formula:

```text
L = √(a² + b²)
```

where `a` and `b` are the horizontal and vertical bay spans the member covers (e.g. 3000 × 2000 → 3605.55).

Tune geometry helpers in `backend/main.py`:
- `length_from_parallel_dimension()` — orthogonal beams
- `triangle_diagonal_length()` / `diagonal_length_from_bay()` — braces

## Editing regex patterns

Open `backend/main.py` and find the section marked:

```text
# REGEX PATTERNS — edit these to match your drawing conventions
```

Tune `BEAM_MARK_RE`, `COLUMN_MARK_RE`, `SECTION_SIZE_RE`, `ELEVATION_RE`, etc. for your office standards (ISMB, W-shapes, UB/UC, etc.).


## UI flow

1. **Upload** — drag & drop PDF (shows name + size)  
2. **Page select** (only if >5 pages) — plan vs elevation pages  
3. **Loading** — progress bar: “AI is reading your drawing…”  
4. **Results** — table preview by tab + **Download Excel**

## Notes

- Large scanned sheets may need higher OCR DPI in `ocr_page_image()` (default 200).  
- Without `OPENAI_API_KEY`, extraction relies on regex + OCR only.  
- Temporary PDFs are stored under the system temp directory and removed after extraction.
