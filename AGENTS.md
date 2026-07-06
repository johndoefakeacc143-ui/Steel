# AGENTS.md

## Cursor Cloud specific instructions

Steel Structure PDF BOM Reader — a single Python CLI (`app.py`) that extracts a
Bill of Materials from steel-structure PDF drawings and writes an Excel file to
`output/<name>_BOM.xlsx`. See `README.md` for full usage and column reference.

### Environment
- Python dependencies are installed into a project virtualenv at `.venv`.
  Always run the app/tools with `.venv/bin/python` (the venv is not auto-activated).
- System packages `tesseract-ocr` (OCR engine for image-based drawings) and
  `python3-venv` are preinstalled in the VM snapshot; they are not part of the
  update script. Tesseract is only exercised for pages with little embedded text.

### Running
- Process every PDF in `input/`: `.venv/bin/python app.py`
- Process one file: `.venv/bin/python app.py --pdf ABC.pdf`
- Skip OCR fallback: `.venv/bin/python app.py --no-ocr`
- Do NOT use `-v`/`--verbose` for routine runs: it enables pdfminer DEBUG logging
  that floods the terminal with thousands of low-value lines.

### Testing / linting
- This repo has no automated test suite and no configured linter/formatter.
  A basic static check is `.venv/bin/python -m py_compile app.py src/*.py`.
- Generated `output/*.xlsx` files are gitignored (except `output/.gitkeep`).
