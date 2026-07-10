"""
NEXT-GEN PRO — Automated Structural Takeoff Backend
FastAPI + Google Gemini + pdf2image + pandas
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional

import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
from pydantic import BaseModel, Field, field_validator

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("nextgen-pro")

# ---------------------------------------------------------------------------
# Paths & configuration
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "generated_exports"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Load GEMINI_API_KEY from .env in the project root.
# Shell exports still win when override=False; .env fills in when unset.
ENV_FILE_PATH = BASE_DIR / ".env"
load_dotenv(ENV_FILE_PATH, override=False)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL_NAME = "gemini-2.5-flash"
PDF_DPI = 200
MAX_PAGES_TO_ANALYZE = 10
SAFE_FILENAME_PATTERN = re.compile(r"^takeoff_[a-zA-Z0-9_\-]+\.xlsx$")

PLACEHOLDER_API_KEYS = {
    "",
    "your_google_gemini_api_key_here",
    "your_key",
    "your_real_key",
    "your_api_key_here",
}

MISSING_API_KEY_DETAIL = (
    "GEMINI_API_KEY is not configured. "
    "1) Get a key at https://aistudio.google.com/apikey  "
    "2) Copy .env.example to .env  "
    "3) Set GEMINI_API_KEY=your_real_key in .env  "
    "4) Restart the server (or run start.bat on Windows)."
)
# ---------------------------------------------------------------------------
# Pydantic schemas — strict Structured JSON Output for Gemini
# ---------------------------------------------------------------------------


class StructuralItem(BaseModel):
    """A single extracted structural component from the drawing."""

    component_mark: str = Field(
        ...,
        description="Component mark or designation, e.g. PB1, PBA, ISMB 300, Grid Span",
    )
    member_type: str = Field(
        ...,
        description="Member type, e.g. Beam, Channel, Bracing, Plate, Column",
    )
    size_or_length: str = Field(
        ...,
        description="Dimensions in mm or structural profile size",
    )
    quantity_count: int = Field(
        ...,
        description="Total count or frequency of this component",
        ge=0,
    )
    engineering_remarks: str = Field(
        ...,
        description="Elevations, placement notes, grid references, or other remarks",
    )

    @field_validator("quantity_count", mode="before")
    @classmethod
    def coerce_quantity(cls, value: object) -> int:
        if value is None or value == "":
            return 0
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            digits = re.findall(r"\d+", value.replace(",", ""))
            return int(digits[0]) if digits else 0
        return 0


class TakeoffResult(BaseModel):
    """Root schema: list of structural takeoff items extracted from the drawing."""

    items: List[StructuralItem] = Field(
        ...,
        description="List of all structural components identified on the drawing",
    )


# ---------------------------------------------------------------------------
# FastAPI application
# ---------------------------------------------------------------------------
app = FastAPI(
    title="NEXT-GEN PRO — Automated Structural Takeoff",
    description="AI-powered structural drawing takeoff using Google Gemini vision",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Gemini helpers (google-genai SDK)
# ---------------------------------------------------------------------------


def get_gemini_api_key() -> str:
    """
    Load GEMINI_API_KEY from the process environment and project .env file.

    Reloads .env on each call so editing .env + uvicorn --reload (or a restart)
    picks up the key. Placeholder values from .env.example are treated as unset.
    """
    global GEMINI_API_KEY

    # Re-read .env so a newly created/edited file is visible after reload.
    if ENV_FILE_PATH.exists():
        load_dotenv(ENV_FILE_PATH, override=True)
    else:
        load_dotenv(ENV_FILE_PATH, override=False)

    raw = os.getenv("GEMINI_API_KEY", "").strip().strip('"').strip("'")
    if raw.lower() in PLACEHOLDER_API_KEYS:
        GEMINI_API_KEY = ""
    else:
        GEMINI_API_KEY = raw
    return GEMINI_API_KEY


def build_gemini_client() -> genai.Client:
    """Create a google-genai Client configured with the API key from .env / env."""
    api_key = get_gemini_api_key()
    if not api_key:
        env_exists = ENV_FILE_PATH.exists()
        logger.error(
            "Gemini API key missing. .env exists=%s path=%s",
            env_exists,
            ENV_FILE_PATH,
        )
        raise HTTPException(status_code=503, detail=MISSING_API_KEY_DETAIL)
    client = genai.Client(api_key=api_key)
    logger.info(
        "Gemini client ready (model=%s, key_loaded=True, key_length=%s).",
        GEMINI_MODEL_NAME,
        len(api_key),
    )
    return client


TAKEOFF_PROMPT = """You are an elite Structural Engineering Estimator performing an automated steel/structural takeoff.

Analyze the attached structural engineering drawing image carefully and extract EVERY identifiable structural component.

For each component, provide:
- component_mark: The mark/tag shown on the drawing (e.g. PB1, PBA, B1, C2, ISMB 300, Grid A-B, PL10). If no mark is visible, invent a logical sequential mark based on member type.
- member_type: Classify as one of: Beam, Column, Channel, Bracing, Plate, Purlin, Rafter, Truss, Connection, Anchor Bolt, Other (use the closest match).
- size_or_length: Profile size and/or length in mm (e.g. "ISMB 300 x 6000", "200x100x10", "PL 10mm x 300x200"). Use mm units.
- quantity_count: Integer count of how many times this identical component appears (or the quantity noted on the drawing).
- engineering_remarks: Elevations, grid locations, levels, welding notes, or any placement/context notes visible on the drawing. Use "—" if none.

Rules:
1. Be thorough — capture beams, columns, bracing, plates, channels, and any tabulated schedules.
2. Prefer reading schedule/BOM tables on the drawing when present; they override scattered callouts.
3. Do not invent quantities that contradict the drawing; if uncertain, set quantity_count to 1 and note uncertainty in engineering_remarks.
4. Return ONLY valid JSON matching the required schema (a root object with an "items" array).
"""


def image_to_jpeg_bytes(pil_image) -> bytes:
    """Convert a PIL Image to JPEG bytes for Gemini upload."""
    buffer = io.BytesIO()
    rgb_image = pil_image.convert("RGB")
    rgb_image.save(buffer, format="JPEG", quality=92)
    return buffer.getvalue()


def analyze_drawing_pages(jpeg_pages: List[bytes]) -> TakeoffResult:
    """Send drawing page images to Gemini and parse structured takeoff JSON."""
    client = build_gemini_client()

    content_parts: list = [TAKEOFF_PROMPT]
    for index, jpeg_bytes in enumerate(jpeg_pages, start=1):
        content_parts.append(f"\n--- Drawing Page {index} of {len(jpeg_pages)} ---")
        content_parts.append(
            types.Part.from_bytes(data=jpeg_bytes, mime_type="image/jpeg")
        )
        logger.info("Attached page %s (%s bytes) for Gemini analysis.", index, len(jpeg_bytes))

    generation_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=TakeoffResult,
        temperature=0.1,
    )

    logger.info("Sending %s page(s) to Gemini for deep structural scan...", len(jpeg_pages))
    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL_NAME,
            contents=content_parts,
            config=generation_config,
        )
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Gemini generate_content failed.")
        raise HTTPException(
            status_code=502,
            detail=f"Gemini API request failed: {exc}",
        ) from exc

    # Prefer the SDK's validated Pydantic parse when available.
    if getattr(response, "parsed", None) is not None:
        parsed_obj = response.parsed
        if isinstance(parsed_obj, TakeoffResult):
            result = parsed_obj
        else:
            result = TakeoffResult.model_validate(parsed_obj)
        logger.info("Parsed %s structural items from Gemini response.parsed.", len(result.items))
        return result

    raw_text = (response.text or "").strip()
    if not raw_text:
        logger.error("Gemini returned an empty response.")
        raise HTTPException(
            status_code=502,
            detail="Gemini returned an empty response. Please retry with a clearer drawing PDF.",
        )

    logger.info("Gemini raw JSON length: %s characters", len(raw_text))

    try:
        parsed = json.loads(raw_text)
        result = TakeoffResult.model_validate(parsed)
    except (json.JSONDecodeError, ValueError) as exc:
        logger.exception("Failed to parse Gemini structured output.")
        raise HTTPException(
            status_code=502,
            detail=f"Failed to parse Gemini structured output: {exc}",
        ) from exc

    logger.info("Parsed %s structural items from Gemini response.", len(result.items))
    return result


def takeoff_to_excel(result: TakeoffResult, source_filename: str) -> Path:
    """Format takeoff items into a pandas DataFrame and save as .xlsx."""
    rows = [
        {
            "Component Mark": item.component_mark,
            "Member Type": item.member_type,
            "Size / Length": item.size_or_length,
            "Quantity": item.quantity_count,
            "Engineering Remarks": item.engineering_remarks,
        }
        for item in result.items
    ]

    if not rows:
        rows = [
            {
                "Component Mark": "—",
                "Member Type": "—",
                "Size / Length": "—",
                "Quantity": 0,
                "Engineering Remarks": "No structural components detected on the uploaded drawing.",
            }
        ]

    dataframe = pd.DataFrame(rows)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    unique_id = uuid.uuid4().hex[:8]
    file_name = f"takeoff_{timestamp}_{unique_id}.xlsx"
    output_path = OUTPUT_DIR / file_name

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Structural Takeoff")
        meta_df = pd.DataFrame(
            [
                {"Field": "Source Drawing", "Value": source_filename},
                {"Field": "Generated At (UTC)", "Value": datetime.now(timezone.utc).isoformat()},
                {"Field": "Model", "Value": GEMINI_MODEL_NAME},
                {"Field": "Item Count", "Value": len(result.items)},
            ]
        )
        meta_df.to_excel(writer, index=False, sheet_name="Metadata")

    logger.info("Excel takeoff saved to: %s", output_path)
    return output_path


# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
async def serve_frontend() -> HTMLResponse:
    """Serve the main dashboard UI."""
    index_path = BASE_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="index.html not found.")
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))


@app.get("/health")
async def health_check() -> dict:
    """Simple health probe for monitoring and frontend setup checks."""
    key_configured = bool(get_gemini_api_key())
    return {
        "status": "ok" if key_configured else "setup_required",
        "service": "NEXT-GEN PRO Structural Takeoff",
        "model": GEMINI_MODEL_NAME,
        "gemini_key_configured": key_configured,
        "env_file_found": ENV_FILE_PATH.exists(),
        "setup_hint": None if key_configured else MISSING_API_KEY_DETAIL,
    }


@app.post("/upload-drawing/")
async def upload_drawing(file: UploadFile = File(...)) -> dict:
    """
    Accept a structural drawing PDF, convert pages to JPEG at 200 DPI,
    run Gemini vision takeoff, and export results to Excel.
    """
    logger.info("Received upload: filename=%s content_type=%s", file.filename, file.content_type)

    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided.")

    lower_name = file.filename.lower()
    if not lower_name.endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Only PDF drawings are supported. Please upload a .pdf file.",
        )

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Uploaded PDF is empty.")

    logger.info("PDF size: %s bytes. Converting pages to JPEG at %s DPI...", len(pdf_bytes), PDF_DPI)

    try:
        pages = convert_from_bytes(pdf_bytes, dpi=PDF_DPI, fmt="jpeg")
    except Exception as exc:
        logger.exception("pdf2image conversion failed.")
        raise HTTPException(
            status_code=422,
            detail=(
                "Failed to convert PDF to images. Ensure Poppler is installed "
                f"and the PDF is valid. Details: {exc}"
            ),
        ) from exc

    if not pages:
        raise HTTPException(status_code=422, detail="PDF contains no readable pages.")

    if len(pages) > MAX_PAGES_TO_ANALYZE:
        logger.warning(
            "PDF has %s pages; analyzing first %s only.",
            len(pages),
            MAX_PAGES_TO_ANALYZE,
        )
        pages = pages[:MAX_PAGES_TO_ANALYZE]

    logger.info("Converted %s page(s). Preparing JPEG payloads...", len(pages))
    jpeg_pages = [image_to_jpeg_bytes(page) for page in pages]

    takeoff_result = analyze_drawing_pages(jpeg_pages)
    excel_path = takeoff_to_excel(takeoff_result, source_filename=file.filename)

    preview_items = [
        {
            "component_mark": item.component_mark,
            "member_type": item.member_type,
            "size_or_length": item.size_or_length,
            "quantity_count": item.quantity_count,
            "engineering_remarks": item.engineering_remarks,
        }
        for item in takeoff_result.items[:5]
    ]

    response_payload = {
        "success": True,
        "message": "AI deep scan complete. Structural takeoff matrix generated.",
        "file_name": excel_path.name,
        "download_url": f"/download/{excel_path.name}",
        "total_items": len(takeoff_result.items),
        "pages_analyzed": len(jpeg_pages),
        "preview": preview_items,
    }
    logger.info(
        "Upload pipeline complete: %s items → %s",
        response_payload["total_items"],
        excel_path.name,
    )
    return response_payload


@app.get("/download/{file_name}")
async def download_excel(file_name: str) -> FileResponse:
    """Safely download a previously generated takeoff Excel spreadsheet."""
    if not SAFE_FILENAME_PATTERN.match(file_name):
        logger.warning("Rejected unsafe download request: %s", file_name)
        raise HTTPException(status_code=400, detail="Invalid file name.")

    file_path = (OUTPUT_DIR / file_name).resolve()
    if not str(file_path).startswith(str(OUTPUT_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Invalid file path.")

    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    logger.info("Serving download: %s", file_path.name)
    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_path.name,
    )


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn

    key_ok = bool(get_gemini_api_key())
    logger.info(
        "Starting NEXT-GEN PRO on http://0.0.0.0:8000 | .env=%s | GEMINI_API_KEY=%s",
        "found" if ENV_FILE_PATH.exists() else "missing",
        "configured" if key_ok else "NOT SET",
    )
    if not key_ok:
        logger.warning(
            "Create .env from .env.example, set GEMINI_API_KEY, then restart. "
            "On Windows you can run start.bat"
        )
    uvicorn.run("backend:app", host="0.0.0.0", port=8000, reload=True)
