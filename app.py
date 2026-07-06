#!/usr/bin/env python3
"""Read steel structure PDF drawings and export a Bill of Materials to Excel."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from src.bom_parser import BOMParser
from src.excel_exporter import ExcelExporter
from src.ocr import OCREngine
from src.pdf_reader import PDFReader

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


def find_pdf_files(input_dir: Path, pdf_name: str | None) -> list[Path]:
    if pdf_name:
        pdf_path = input_dir / pdf_name
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        return [pdf_path]

    pdfs = sorted(input_dir.glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError(f"No PDF files found in {input_dir}")
    return pdfs


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


def process_pdf(pdf_path: Path, output_dir: Path, use_ocr: bool) -> Path:
    logging.info("Processing %s", pdf_path.name)
    reader = PDFReader(pdf_path)
    pages = reader.read()

    if use_ocr:
        ocr = OCREngine()
        enrich_pages_with_ocr(pages, ocr)

    parser = BOMParser()
    bom_df = parser.parse_pages(pages)

    output_name = f"{pdf_path.stem}_BOM.xlsx"
    output_path = output_dir / output_name
    exporter = ExcelExporter()
    exporter.export(bom_df, output_path)

    logging.info(
        "Exported %s items to %s",
        len(bom_df),
        output_path,
    )
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract Bill of Materials from steel structure PDF drawings.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=INPUT_DIR,
        help=f"Directory containing PDF files (default: {INPUT_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_DIR,
        help=f"Directory for Excel output (default: {OUTPUT_DIR})",
    )
    parser.add_argument(
        "--pdf",
        type=str,
        default=None,
        help="Process a single PDF file from the input directory",
    )
    parser.add_argument(
        "--no-ocr",
        action="store_true",
        help="Disable OCR fallback for image-based pages",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable debug logging",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    setup_logging(args.verbose)

    try:
        pdf_files = find_pdf_files(args.input_dir, args.pdf)
    except FileNotFoundError as exc:
        logging.error("%s", exc)
        return 1

    args.output_dir.mkdir(parents=True, exist_ok=True)
    use_ocr = not args.no_ocr
    output_files: list[Path] = []

    for pdf_path in pdf_files:
        try:
            output_files.append(process_pdf(pdf_path, args.output_dir, use_ocr))
        except Exception:
            logging.exception("Failed to process %s", pdf_path.name)
            return 1

    for path in output_files:
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
