# SteelDraw AI Extractor

Full-stack app that reads steel structure engineering PDF drawings like a
detailer/BIM modeler and extracts **Beams**, **Columns**, **Bracing**, and
**Base Plates** into an Excel workbook.

## Features

- Upload PDFs up to **500 MB**, processed **page by page**
- Detects **digital vs scanned** pages; scanned pages use OpenCV + Tesseract OCR
- Optional **Camelot** table extraction for member schedules
- **Regex + Gemini / OpenAI vision (LangChain)** — page IMAGE is sent so the model
  reads bay dimensions (e.g. BR1 beside 3000×1500 → √(a²+b²)) from THIS drawing
- No hardcoded project sizes — next drawing can have different bay legs
- **Page selection** when the drawing has more than 5 pages:
  - Plan pages → beams & bracing
  - Elevation pages → columns & elevation / base plates
- Drawings with **1–5 pages** scan automatically (no page prompt)
- Excel download with 4 fabrication sheets + Summary

### Excel sheets

| Sheet | Columns |
|-------|---------|
| Beams | Mark \| Length_mm \| Quantity |
| Columns | Mark \| Height_mm \| Quantity |
| Bracing | Mark \| Length_mm \| Quantity |
| BasePlates | Mark \| Plate_Size_mm \| Weight |
| Summary | Total Beams/Columns/Bracing/BasePlates, Min/Max Elevation |

## Folder structure

```
/backend
  .env                 # GEMINI_API_KEY / OPENAI_API_KEY
  main.py              # FastAPI app + extraction pipeline (edit REGEX here)
  requirements.txt
/frontend
  index.html
  package.json
  vite.config.js
  tailwind.config.js
  src/App.jsx          # Upload → Page select → Loading → Results
  src/main.jsx
  src/index.css
.env.example
README.md
```

## Prerequisites

- Python 3.10+
- Node.js 18+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- Optional (Camelot lattice tables): Ghostscript
- Gemini API key from [Google AI Studio](https://aistudio.google.com/apikey)
  (or OpenAI key as fallback)

### Install Tesseract

```bash
# Ubuntu / Debian
sudo apt-get update && sudo apt-get install -y tesseract-ocr ghostscript

# macOS
brew install tesseract ghostscript
```

## Setup

### 1. Environment variables

```bash
# Windows (Command Prompt)
copy .env.example .env

# Windows (PowerShell)
Copy-Item .env.example .env

# macOS / Linux
cp .env.example .env
```

Edit the project-root `.env` (and/or `backend/.env`):

```env
GEMINI_API_KEY=AIza-your-real-key
GEMINI_MODEL=gemini-2.0-flash

# Optional fallback:
# OPENAI_API_KEY=sk-proj-your-real-key
# OPENAI_MODEL=gpt-4o-mini
```

No quotes. One line per key. Then restart uvicorn.

Confirm AI is on: [http://127.0.0.1:8000/api/health](http://127.0.0.1:8000/api/health)
→ `"ai_configured": true`, `"ai_provider": "gemini"`.

On startup the backend prints `[startup] AI READY` or `AI OFF`.
Regex/OCR extraction works without an API key; Gemini/OpenAI improves mark
association and the Summary narrative.

### 2. Backend

Use **Python 3.10–3.12** (3.11 or 3.12 recommended on Windows).

```bash
cd backend
python -m venv .venv

# Windows (Command Prompt / PowerShell):
.venv\Scripts\activate

# macOS / Linux:
source .venv/bin/activate

python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

If `numpy` fails with `metadata-generation-failed` on Windows:

```bat
python -m pip install --upgrade pip setuptools wheel
pip install numpy
pip install -r requirements.txt
```

Or install a prebuilt wheel first, then the rest:

```bat
pip install "numpy>=1.26.4,<2.3"
pip install -r requirements.txt
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
| `plan_pages` | e.g. `1,2` or `1-3` — beams & bracing |
| `elevation_pages` | e.g. `4,5` — columns & base plates |
| `confirm` | `true` after the user picks pages (>5 page drawings) |

**Flow**

1. Upload PDF.
2. If `page_count > 5` → response `{ needs_page_selection: true, job_id, ... }`.
   UI asks: *From which page of plan…?* / *From which page of elevation…?*
3. If `page_count ≤ 5` → extract immediately on all pages.
4. Success response includes `preview` tables + `excel_base64` for Download Excel.

### `GET /api/health`

Returns AI provider status (`gemini` preferred, else `openai`).

### `POST /api/download`

Same pipeline; streams the `.xlsx` file directly (for API clients).

## AI prompt (senior detailer)

The backend system prompt instructs the model as a Senior Steel Structure
Detailer / BIM Modeler (15 years) to extract exactly the 4 tables above, use
`N/A` when unreadable, and never guess for fabrication. Edit
`AI_SYSTEM_PROMPT` in `backend/main.py` to tune behaviour.

## Editing regex

All mark / section / dimension patterns live near the top of `backend/main.py`
under **REGEX PATTERNS** — search for `BEAM_MARK_RE`, `COLUMN_MARK_RE`,
`BRACING_MARK_RE`, `BASE_PLATE_MARK_RE`. Comments explain each pattern so you
can adapt them to your drawing conventions.

## UI screens

1. **Upload** — drag & drop PDF, show file name + size
2. **Page select** (only if >5 pages) — plan vs elevation page numbers
3. **Loading** — progress bar: “AI is reading your drawing…”
4. **Results** — table preview tabs + Download Excel

## License

Internal / project use.
