# NEXT-GEN PRO — Automated Structural Takeoff

AI-powered structural drawing takeoff application. Upload a PDF drawing, run a Gemini vision deep scan, and download a clean Excel bill of materials.

## Stack

- **Backend:** FastAPI + Uvicorn
- **Vision AI:** Google Gemini (`gemini-2.5-flash`) via the official `google-genai` SDK with structured JSON output
- **PDF → Image:** `pdf2image` at 200 DPI (requires Poppler)
- **Export:** pandas → Excel (`.xlsx`)
- **Frontend:** HTML5 + Tailwind CSS + vanilla JavaScript (dark slate dashboard)

## Project structure

```
.
├── backend.py          # FastAPI application
├── index.html          # Dashboard UI
├── start.bat           # Windows launcher (loads .env + starts server)
├── requirements.txt    # Python dependencies
├── .env.example        # API key template (copy to .env)
└── generated_exports/  # Generated Excel files (created at runtime)
```

## Prerequisites

1. **Python 3.10+**
2. **Poppler** (required by `pdf2image`)
   - Ubuntu/Debian: `sudo apt-get install -y poppler-utils`
   - Windows: install Poppler and add its `bin` folder to PATH
   - macOS: `brew install poppler`
3. **Google Gemini API key** from [Google AI Studio](https://aistudio.google.com/apikey)

## Setup (enable the Gemini API via `.env`)

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

### 1. Create `.env`

```bat
REM Windows
copy .env.example .env
notepad .env
```

```bash
# macOS / Linux
cp .env.example .env
nano .env
```

### 2. Put your real key in `.env`

```env
GEMINI_API_KEY=AIzaSyYourRealKeyHere
```

Rules that matter:

- No quotes around the key
- No spaces around `=`
- Do not leave the placeholder `your_google_gemini_api_key_here`

### 3. Start the server

**Windows (recommended):** double-click or run:

```bat
start.bat
```

`start.bat` loads `GEMINI_API_KEY` from `.env` and starts Uvicorn.

**Or manually:**

```bash
uvicorn backend:app --host 0.0.0.0 --port 8000 --reload
```

Open **http://127.0.0.1:8000**

If the amber **Setup Required** banner still appears, your `.env` key is missing/placeholder — fix `.env`, then **fully stop** the server (Ctrl+C) and start again.

Check status:

```bash
curl http://127.0.0.1:8000/health
```

You want `"gemini_key_configured": true` and `"env_file_found": true`.

## API

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/` | Serves the dashboard UI |
| `GET` | `/health` | Health check + whether `.env` key loaded |
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
- Never commit your real `.env` file (it is gitignored).
