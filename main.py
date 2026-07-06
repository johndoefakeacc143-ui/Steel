#!/usr/bin/env python3
"""Anu-AI - read steel structure drawings from ./input and build a BOM.

Reads every PDF/DWG/DXF file in the input folder, extracts steel members
(ID, type, section, length, quantity, elevation, grid, weight), converts
lengths to millimetres, computes missing weights from the section catalogue,
and writes ./output/BOM.xlsx with a Detailed_BOM sheet and a Summary sheet.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from src.bom import build_detailed_bom, build_summary, build_type_counts
from src.extractor import MemberExtractor
from src.readers import SUPPORTED_SUFFIXES, read_file

PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"

logger = logging.getLogger("anu_ai")


def setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    # pdfminer is extremely chatty at DEBUG; keep it quiet.
    logging.getLogger("pdfminer").setLevel(logging.WARNING)


def find_drawings(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")
    files = sorted(
        p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES
    )
    return files


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Anu-AI steel drawing BOM generator")
    parser.add_argument("--input-dir", type=Path, default=INPUT_DIR)
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR / "BOM.xlsx")
    parser.add_argument("--no-ocr", action="store_true", help="Disable OCR for scanned PDFs")
    parser.add_argument(
        "--engine",
        choices=["classic", "ai"],
        default="classic",
        help="Extraction engine: 'classic' (pdfplumber/OCR) or 'ai' (vision LLM)",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    setup_logging(args.verbose)

    try:
        drawings = find_drawings(args.input_dir)
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        return 1

    if not drawings:
        logger.error("No PDF/DWG/DXF files found in %s", args.input_dir)
        return 1

    ai_reader = None
    if args.engine == "ai":
        from src.ai_reader import AIConfigError, VisionAIReader

        ai_reader = VisionAIReader()
        if not ai_reader.available():
            logger.error(
                "AI engine selected but no API key found. "
                "Set ANU_AI_API_KEY (or OPENAI_API_KEY). See README."
            )
            return 2
        logger.info("Using AI vision engine: model=%s", ai_reader.model)

    extractor = MemberExtractor()
    members = []
    for path in drawings:
        try:
            logger.info("Reading %s", path.name)
            if ai_reader is not None:
                if path.suffix.lower() != ".pdf":
                    logger.warning("  AI engine supports PDF only; skipping %s", path.name)
                    continue
                found = ai_reader.read(path)
            else:
                raw = read_file(path, use_ocr=not args.no_ocr)
                found = extractor.extract(raw)
            logger.info("  extracted %d members from %s", len(found), path.name)
            members.extend(found)
        except Exception:  # noqa: BLE001 - one bad file must not abort the run
            logger.exception("Failed to process %s", path.name)

    detailed = build_detailed_bom(members)
    summary = build_summary(members)
    type_counts = build_type_counts(members)

    from src.excel_writer import write_bom

    output_path = write_bom(detailed, summary, type_counts, args.output)

    logger.info("Total members: %d", len(members))
    for _, row in type_counts.iterrows():
        logger.info("  %-10s : %s", row["Member_Type"], row["Total_Count"])
    logger.info("BOM written to %s", output_path)
    print(output_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
