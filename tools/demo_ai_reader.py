#!/usr/bin/env python3
"""Offline demonstration of the AI vision reader pipeline.

Stubs ONLY the network call to the vision model with a canned JSON response, so
everything else runs for real: PDF rendering + contrast enhancement + base64
encoding, JSON parsing, member finalisation (weight calc), and the two-sheet BOM
build. Proves the AI engine plumbing without needing an API key.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.ai_reader import VisionAIReader
from src.bom import build_detailed_bom, build_summary, build_type_counts

# A realistic model response for a GA drawing (as the vision LLM would return).
FAKE_RESPONSE = json.dumps(
    {
        "members": [
            {"member_id": "B7", "member_type": "Beam", "section_size": "ISMB 300",
             "length_mm": 2000, "quantity": 76, "elevation_m": 106.0,
             "grid_location": "A-1", "weight_kg": None},
            {"member_id": "B8", "member_type": "Beam", "section_size": "ISMB 400",
             "length_mm": 2000, "quantity": 24, "elevation_m": 106.0,
             "grid_location": "B-2", "weight_kg": None},
            {"member_id": "BR1", "member_type": "Bracing", "section_size": "ISMC 150",
             "length_mm": 3000, "quantity": 12, "elevation_m": None,
             "grid_location": "C-3", "weight_kg": None},
            {"member_id": "BP1", "member_type": "Plate", "section_size": "PL 400x20",
             "length_mm": 400, "quantity": 11, "elevation_m": 100.3,
             "grid_location": None, "weight_kg": None},
        ]
    }
)


def main() -> None:
    pdf = Path("input/SAMPLE_STRUCTURE.pdf")
    reader = VisionAIReader()
    reader.api_key = "test-key"  # bypass availability check

    # Confirm real rendering works (PDF -> enhanced PNG -> base64).
    images = reader._render_pages(pdf)
    print(f"Rendered {len(images)} page image(s); first is {len(images[0])} base64 chars")

    # Stub only the network call.
    reader._call_model = lambda image_b64: FAKE_RESPONSE  # type: ignore[assignment]

    members = reader.read(pdf)
    print(f"AI reader produced {len(members)} members\n")

    detailed = build_detailed_bom(members)
    print("=== Detailed_BOM ===")
    print(detailed.to_string(index=False))
    print("\n=== Summary (per Section_Size) ===")
    print(build_summary(members).to_string(index=False))
    print("\n=== Member Type Counts ===")
    print(build_type_counts(members).to_string(index=False))


if __name__ == "__main__":
    main()
