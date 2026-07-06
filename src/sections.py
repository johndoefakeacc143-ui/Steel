"""Steel section catalogue and unit-weight helpers.

Provides nominal unit weights (kg/m) for common Indian standard sections and
helpers to compute weights for rolled hollow sections (RHS) and plates (PL)
from their designation when they are not tabulated.
"""

from __future__ import annotations

import re

STEEL_DENSITY_KG_PER_MM3 = 7.85e-6  # 7850 kg/m^3 expressed per mm^3

# Nominal unit weights (kg/m) for common ISMB (I-beams) by nominal depth (mm).
ISMB_WEIGHTS: dict[int, float] = {
    100: 11.5,
    125: 13.0,
    150: 15.0,
    175: 19.3,
    200: 25.4,
    225: 31.2,
    250: 37.3,
    300: 44.2,
    350: 52.4,
    400: 61.6,
    450: 72.4,
    500: 86.9,
    550: 103.7,
    600: 122.6,
}

# Nominal unit weights (kg/m) for common ISMC (channels) by nominal depth (mm).
ISMC_WEIGHTS: dict[int, float] = {
    75: 7.14,
    100: 9.56,
    125: 13.1,
    150: 16.4,
    175: 19.1,
    200: 22.1,
    225: 25.9,
    250: 30.4,
    300: 35.8,
    400: 49.4,
}

# Section-type inference from a designation prefix.
SECTION_TYPE_ALIASES = {
    "ISMB": "ISMB",
    "MB": "ISMB",
    "ISMC": "ISMC",
    "MC": "ISMC",
    "RHS": "RHS",
    "SHS": "RHS",
    "PL": "PL",
    "PLATE": "PL",
    "PLT": "PL",
}

_NOMINAL_RE = re.compile(r"(\d{2,4})")
_RHS_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)"
)
_PL_RE = re.compile(
    r"(?:PL|PLATE|PLT)\s*[.:]?\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)",
    re.IGNORECASE,
)


def normalize_section_type(section: str) -> str | None:
    """Return the canonical section family (ISMB/ISMC/RHS/PL) or None."""
    if not section:
        return None
    upper = section.upper()
    for alias, canonical in SECTION_TYPE_ALIASES.items():
        if re.match(rf"^\s*{alias}\b", upper) or upper.startswith(alias):
            return canonical
    return None


def unit_weight_kg_per_m(section: str) -> float | None:
    """Best-effort unit weight (kg/m) for a section designation.

    Returns None when the weight cannot be determined from the designation.
    Plates (PL) are length-dependent and handled by ``plate_weight_kg``.
    """
    if not section:
        return None

    family = normalize_section_type(section)
    upper = section.upper()

    if family in ("ISMB", "ISMC"):
        match = _NOMINAL_RE.search(upper)
        if not match:
            return None
        nominal = int(match.group(1))
        table = ISMB_WEIGHTS if family == "ISMB" else ISMC_WEIGHTS
        return table.get(nominal)

    if family == "RHS":
        match = _RHS_RE.search(upper)
        if not match:
            return None
        height, width, thickness = (float(g) for g in match.groups())
        # Thin-walled RHS cross-section area (mm^2), ignoring corner radii.
        area = 2.0 * thickness * (height + width - 2.0 * thickness)
        return area * 1000.0 * STEEL_DENSITY_KG_PER_MM3

    return None


def plate_weight_kg(section: str, length_mm: float | None) -> float | None:
    """Total weight (kg) of a plate given ``PL width x thickness`` and length."""
    if not section or length_mm is None:
        return None
    match = _PL_RE.search(section)
    if not match:
        return None
    width_mm, thickness_mm = (float(g) for g in match.groups())
    volume_mm3 = width_mm * thickness_mm * length_mm
    return volume_mm3 * STEEL_DENSITY_KG_PER_MM3


def compute_weight_kg(section: str, length_mm: float | None) -> float | None:
    """Compute per-member weight (kg) from section and length.

    Uses ``Length * unit weight`` for rolled sections and a volumetric
    calculation for plates. Returns None when it cannot be determined.
    """
    if normalize_section_type(section) == "PL":
        return plate_weight_kg(section, length_mm)

    unit = unit_weight_kg_per_m(section)
    if unit is None or length_mm is None:
        return None
    return (length_mm / 1000.0) * unit
