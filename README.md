# Anu-AI

Read steel structure drawings (PDF / DWG / DXF) and generate a Bill of Materials
in Excel.

## Features

- Reads every `*.pdf`, `*.dwg`, `*.dxf` file in `input/`.
- Extracts per steel member: `Member_ID`, `Member_Type` (Beam / Column / Bracing
  / Plate / Girder), `Section_Size` (ISMB / ISMC / RHS / PL), `Length_mm`,
  `Quantity`, `Elevation_m`, `Grid_Location`, `Weight_kg`.
- Lengths given in metres are auto-converted to millimetres.
- Missing weight is calculated as `Length × section unit-weight` (plates by volume).
- Missing values are written as `NA`.
- Text PDFs use `pdfplumber`; scanned PDFs fall back to OpenCV + Tesseract OCR;
  tables and free text are both parsed.
- DWG requires the ODA File Converter (via `ezdxf`'s `odafc` add-on); DXF is read
  directly.

## Setup

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
# System dependency for scanned PDFs:
sudo apt-get install tesseract-ocr
```

## Usage

1. Put drawing files in `input/`.
2. Run:

```bash
python main.py            # process input/ -> output/BOM.xlsx (classic engine)
python main.py --no-ocr   # skip OCR
python main.py -v         # verbose logging
```

## Engines

Two extraction engines are available via `--engine`:

- `classic` (default) — `pdfplumber` for text/tables, with OpenCV + Tesseract OCR
  fallback for scanned pages.
- `ai` — a **vision LLM** reads the drawing image directly. Unlike OCR, it copes
  with faint, low-contrast, and rotated dimension text and returns structured
  members. Recommended for General Arrangement drawings.

```bash
python main.py --engine ai
```

The AI engine is provider-agnostic (OpenAI-compatible Chat Completions), so it
works with OpenAI, OpenRouter, Azure, or a local server (vLLM / Ollama /
llama.cpp). Configure via environment variables (no key is stored in the repo):

```bash
export ANU_AI_API_KEY=sk-...            # or OPENAI_API_KEY
export ANU_AI_MODEL=gpt-4o-mini         # any vision-capable model
export ANU_AI_BASE_URL=https://api.openai.com/v1   # or a local/compatible endpoint
```

To try the pipeline without a key (uses a canned model response):

```bash
python tools/demo_ai_reader.py
```

## Output — `output/BOM.xlsx`

- **Detailed_BOM**: one row per member with all extracted fields.
- **Summary**: total quantity, length and weight per `Section_Size`, plus a
  count of members by `Member_Type`.

## Project layout

```
├── main.py                 # entry point
├── input/                  # drawings (PDF/DWG/DXF)
├── output/                 # BOM.xlsx
├── requirements.txt
├── tools/generate_sample.py  # builds a sample member-schedule PDF
└── src/
    ├── readers.py          # PDF (pdfplumber + OCR) and DWG/DXF (ezdxf) readers
    ├── ai_reader.py        # vision-LLM engine (reads drawings without OCR)
    ├── sections.py         # ISMB/ISMC/RHS/PL unit-weight catalogue
    ├── extractor.py        # table + text member extraction
    ├── bom.py              # detailed + summary tables
    └── excel_writer.py     # 2-sheet Excel output
```
