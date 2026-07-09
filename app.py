#!/usr/bin/env python3
"""Read steel structure PDF drawings and export a Bill of Materials to Excel."""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

from src.bom_parser import BOMParser
from src.dimension_ocr import DimensionEstimator
from src.excel_exporter import ExcelExporter
from src.input_reader import SUPPORTED_SUFFIXES, read_drawing
from src.ocr import OCREngine

PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"

MIN_TEXT_LENGTH = 50


def setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    # These libraries emit thousands of DEBUG lines; keep them quiet so progress
    # logging stays readable even with -v.
    for noisy in ("pdfminer", "pdfplumber", "PIL"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def find_inputs(input_dir: Path, single_name: str | None) -> list[Path]:
    if single_name:
        path = input_dir / single_name
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        return [path]

    files = sorted(
        p
        for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES
    )
    if not files:
        raise FileNotFoundError(
            f"No supported drawings (PDF/image/DWG/DXF) found in {input_dir}"
        )
    return files


def tesseract_available() -> bool:
    """Return True if Tesseract can be called (auto-detecting its location)."""
    from src.tesseract_setup import ensure_tesseract

    return ensure_tesseract()


def enrich_pages_with_ocr(pages: list, ocr: OCREngine) -> None:
    for page in pages:
        if len(page.text.strip()) >= MIN_TEXT_LENGTH:
            continue
        if page.image is None:
            logging.warning("Page %s has no image for OCR", page.page_number)
            continue
        logging.info("Running OCR on page %s", page.page_number)
        ocr_text = ocr.extract_text(page.image)
        page.text = f"{page.text}\n{ocr_text}".strip()


def process_drawing(
    path: Path,
    output_dir: Path,
    use_ocr: bool,
    length_overrides: dict | None = None,
    ai_reader=None,
) -> Path:
    logging.info("Processing %s", path.name)
    parser = BOMParser()

    if ai_reader is not None:
        items = ai_reader.read(path)
        bom_df = parser.dataframe_from_items(items)
    else:
        pages = read_drawing(path)

        mark_lengths: dict = {}
        page_default_lengths: dict = {}
        if use_ocr:
            ocr = OCREngine()
            enrich_pages_with_ocr(pages, ocr)

            # Recover member lengths from drawing dimension lines via OCR.
            estimator = DimensionEstimator()
            mark_lengths, page_default_lengths = estimator.estimate(pages)
            if mark_lengths:
                logging.info(
                    "Estimated lengths for %s marks from dimensions", len(mark_lengths)
                )

        bom_df = parser.parse_pages(
            pages, mark_lengths, page_default_lengths, length_overrides
        )

    output_path = output_dir / f"{path.stem}_BOM.xlsx"
    ExcelExporter().export(bom_df, output_path)
    logging.info("Exported %s items to %s", len(bom_df), output_path)
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract a Bill of Materials from steel structure drawings "
            "(PDF, scanned/image PDF, image files, or DWG/DXF)."
        ),
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=INPUT_DIR,
        help=f"Directory containing drawings (default: {INPUT_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_DIR,
        help=f"Directory for Excel output (default: {OUTPUT_DIR})",
    )
    parser.add_argument(
        "--file",
        "--pdf",
        dest="file",
        type=str,
        default=None,
        help="Process a single file (PDF/image/DWG/DXF) from the input directory",
    )
    parser.add_argument(
        "--engine",
        choices=["classic", "ai"],
        default="classic",
        help="Extraction engine: 'classic' (pdfplumber/OCR) or 'ai' (vision LLM)",
    )
    parser.add_argument(
        "--no-ocr",
        action="store_true",
        help="Disable OCR fallback for image-based pages",
    )
    parser.add_argument(
        "--lengths-file",
        type=Path,
        default=None,
        help=(
            "JSON file of authoritative {mark: length_mm} overrides. "
            f"Defaults to {INPUT_DIR / 'member_lengths.json'} if present."
        ),
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable debug logging",
    )
    return parser


def load_length_overrides(input_dir: Path, lengths_file: Path | None) -> dict[str, float]:
    """Load authoritative per-mark length overrides from JSON, if available."""
    path = lengths_file or (input_dir / "member_lengths.json")
    if not path.exists():
        return {}
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
        # Values are either a scalar length (mm) or a list of length groups.
        overrides = {str(k): v for k, v in data.items()}
        logging.info("Loaded %d length override(s) from %s", len(overrides), path)
        return overrides
    except (OSError, ValueError) as exc:
        logging.warning("Could not read length overrides from %s: %s", path, exc)
        return {}


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    setup_logging(args.verbose)

    try:
        input_files = find_inputs(args.input_dir, args.file)
    except FileNotFoundError as exc:
        logging.error("%s", exc)
        return 1

    ai_reader = None
    if args.engine == "ai":
        from src.ai_reader import VisionAIReader

        ai_reader = VisionAIReader()
        if not ai_reader.available():
            logging.error(
                "AI engine selected but no API key found. "
                "Set ANU_AI_API_KEY (or OPENAI_API_KEY). See README."
            )
            return 2
        logging.info("Using AI vision engine: model=%s", ai_reader.model)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    use_ocr = not args.no_ocr
    if use_ocr and ai_reader is None and not tesseract_available():
        logging.warning(
            "Tesseract OCR is not installed - member lengths read from drawing "
            "dimension lines will be EMPTY (only values from member_lengths.json "
            "will appear). Install it: `sudo apt-get install tesseract-ocr` "
            "(Ubuntu/Debian) or `brew install tesseract` (macOS)."
        )
    length_overrides = load_length_overrides(args.input_dir, args.lengths_file)
    output_files: list[Path] = []

    total = len(input_files)
    logging.info("Found %d drawing(s) to process in %s", total, args.input_dir)
    for index, path in enumerate(input_files, start=1):
        logging.info("=== [%d/%d] %s ===", index, total, path.name)
        try:
            output_files.append(
                process_drawing(
                    path, args.output_dir, use_ocr, length_overrides, ai_reader
                )
            )
        except Exception:
            logging.exception("Failed to process %s", path.name)
            return 1

    logging.info("Done. Wrote %d BOM file(s).", len(output_files))
    for path in output_files:
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
