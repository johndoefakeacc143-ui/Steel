"""
Steel Structure OCR Extractor — FastAPI application

Upload PDF / image drawings (including scanned PDFs with images).
OCR traces beams, columns, and base plates → Excel with sheets:
  Beam | Columns | Base Plate
Each sheet: Member Name, Quantity, Size (+ detail columns).
"""

from __future__ import annotations

import logging
import tempfile
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from steel_ocr.pipeline import process_drawing

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("steel_ocr.app")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
EXPORT_DIR = BASE_DIR / "generated_exports"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

MAX_UPLOAD_BYTES = 200 * 1024 * 1024  # 200 MB
ALLOWED_SUFFIXES = {
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
    ".bmp",
    ".webp",
}

app = FastAPI(
    title="Steel Structure OCR Extractor",
    description=(
        "OCR steel structure PDFs and images; export Beam, Columns, "
        "and Base Plate sheets to Excel."
    ),
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

if STATIC_DIR.exists():
    app.mount("/assets", StaticFiles(directory=STATIC_DIR), name="assets")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    index_path = STATIC_DIR / "index.html"
    if not index_path.exists():
        return HTMLResponse(
            "<h1>Steel OCR</h1><p>static/index.html missing.</p>",
            status_code=500,
        )
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.get("/api/health")
def health() -> dict:
    tesseract_ok = False
    tesseract_version = None
    try:
        import pytesseract

        tesseract_version = str(pytesseract.get_tesseract_version())
        tesseract_ok = True
    except Exception as exc:  # noqa: BLE001
        tesseract_version = str(exc)

    return {
        "status": "ok",
        "tesseract_ok": tesseract_ok,
        "tesseract_version": tesseract_version,
        "accepted_types": sorted(ALLOWED_SUFFIXES),
        "excel_sheets": ["Beam", "Columns", "Base Plate"],
        "sheet_columns": ["Member Name", "Quantity", "Size"],
    }


def _save_upload(upload: UploadFile) -> Path:
    suffix = Path(upload.filename or "drawing.pdf").suffix.lower()
    if suffix not in ALLOWED_SUFFIXES:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported file type '{suffix}'. "
                f"Accepted: {', '.join(sorted(ALLOWED_SUFFIXES))}"
            ),
        )

    dest = Path(tempfile.gettempdir()) / f"steel_ocr_{uuid.uuid4().hex}{suffix}"
    size = 0
    with dest.open("wb") as out:
        while True:
            chunk = upload.file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                out.close()
                dest.unlink(missing_ok=True)
                raise HTTPException(status_code=413, detail="File exceeds 200 MB limit.")
            out.write(chunk)
    return dest


@app.post("/api/extract")
async def extract(file: UploadFile = File(...)) -> JSONResponse:
    """
    Upload a PDF or image, run OCR + member extraction, return preview + Excel (base64).
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="Missing filename.")

    tmp_path: Path | None = None
    try:
        tmp_path = _save_upload(file)
        logger.info("Processing upload %s (%s)", file.filename, tmp_path.suffix)
        result = process_drawing(tmp_path)

        # Persist a copy for /api/download/{id}
        job_id = uuid.uuid4().hex[:12]
        out_name = f"{job_id}_{Path(file.filename).stem}_steel_takeoff.xlsx"
        out_path = EXPORT_DIR / out_name
        out_path.write_bytes(result["excel_bytes"])

        return JSONResponse(
            {
                "ok": True,
                "job_id": job_id,
                "source_filename": file.filename,
                "excel_filename": result["excel_filename"],
                "download_url": f"/api/download/{job_id}",
                "document": result["document"],
                "preview": result["preview"],
                "excel_base64": result["excel_base64"],
            }
        )
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        logger.exception("Extraction failed")
        raise HTTPException(status_code=500, detail=f"Extraction failed: {exc}") from exc
    finally:
        if tmp_path and tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


@app.post("/api/download")
async def download_direct(file: UploadFile = File(...)) -> StreamingResponse:
    """Upload and stream the Excel file directly (API clients / curl)."""
    tmp_path: Path | None = None
    try:
        tmp_path = _save_upload(file)
        result = process_drawing(tmp_path)
        headers = {
            "Content-Disposition": f'attachment; filename="{result["excel_filename"]}"'
        }
        return StreamingResponse(
            iter([result["excel_bytes"]]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        logger.exception("Download extraction failed")
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    finally:
        if tmp_path and tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


@app.get("/api/download/{job_id}")
def download_job(job_id: str) -> FileResponse:
    matches = list(EXPORT_DIR.glob(f"{job_id}_*.xlsx"))
    if not matches:
        raise HTTPException(status_code=404, detail="Export not found or expired.")
    path = matches[0]
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name.split("_", 1)[-1] if "_" in path.name else path.name,
    )


def run() -> None:
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)


if __name__ == "__main__":
    run()
