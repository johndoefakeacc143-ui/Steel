# NEXT-GEN PRO — Automated Structural Takeoff

AI-powered structural drawing takeoff application. Upload a PDF drawing, run a Gemini vision deep scan, and download a clean Excel bill of materials.

## Stack

- **Backend:** FastAPI + Uvicorn
- **Vision AI:** Google Gemini (`gemini-2.5-flash`) with structured JSON output
- **PDF → Image:** `pdf2image` at 200 DPI (requires Poppler)
- **Export:** pandas → Excel (`.xlsx`)
- **Frontend:** HTML5 + Tailwind CSS + vanilla JavaScript (dark slate dashboard)

## Project structure

```
.
├── backend.py          # FastAPI application
├── index.html          # Dashboard UI
├── requirements.txt    # Python dependencies
├── .env.example        # API key template
└── generated_exports/  # Generated Excel files (created at runtime)
```

## Prerequisites

1. **Python 3.10+**
2. **Poppler** (required by `pdf2image`)
   - Ubuntu/Debian: `sudo apt-get install -y poppler-utils`
   - macOS: `brew install poppler`
3. **Google Gemini API key** from [Google AI Studio](https://aistudio.google.com/apikey)

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Required: Google Gemini API key
cp .env.example .env
# Edit .env and set: GEMINI_API_KEY=your_real_key
# Get a free key at https://aistudio.google.com/apikey
```

Alternatively, export the key in your shell (then restart the server):

```bash
export GEMINI_API_KEY="your_api_key_here"
```

> If you see **“GEMINI_API_KEY is not configured”** in the UI, the key is missing from the process environment. Add it via `.env` or `export`, then restart Uvicorn.
## Run

```bash
uvicorn backend:app --host 0.0.0.0 --port 8000 --reload
```

Open **http://127.0.0.1:8000** in your browser.

## API

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/` | Serves the dashboard UI |
| `GET` | `/health` | Health check |
| `POST` | `/upload-drawing/` | Upload PDF → AI takeoff → Excel |
| `GET` | `/download/{file_name}` | Download generated `.xlsx` |

### Upload response shape

```json
{
  "success": true,
  "message": "AI deep scan complete. Structural takeoff matrix generated.",
  "file_name": "takeoff_YYYYMMDD_HHMMSS_abcd1234.xlsx",
  "download_url": "/download/takeoff_....xlsx",
  "total_items": 42,
  "pages_analyzed": 2,
  "preview": [
    {
      "component_mark": "PB1",
      "member_type": "Beam",
      "size_or_length": "ISMB 300 x 6000",
      "quantity_count": 4,
      "engineering_remarks": "Grid A-B · EL +3.150"
    }
  ]
}
```

## Notes

- Multi-page PDFs are supported (first 10 pages analyzed by default).
- CORS is fully enabled for local frontend development.
- Generated Excel files are stored under `generated_exports/` and served only when the filename matches a safe pattern.
