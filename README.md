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
python main.py            # process input/ -> output/BOM.xlsx
python main.py --no-ocr   # skip OCR
python main.py -v         # verbose logging
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
    ├── sections.py         # ISMB/ISMC/RHS/PL unit-weight catalogue
    ├── extractor.py        # table + text member extraction
    ├── bom.py              # detailed + summary tables
    └── excel_writer.py     # 2-sheet Excel output
```
