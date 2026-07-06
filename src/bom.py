"""Assemble detailed and summary Bill of Materials tables."""

from __future__ import annotations

import pandas as pd

from src.extractor import NA, Member

DETAILED_COLUMNS = [
    "Member_ID",
    "Member_Type",
    "Section_Size",
    "Length_mm",
    "Quantity",
    "Elevation_m",
    "Grid_Location",
    "Weight_kg",
    "Source_File",
]


def _na(value) -> object:
    if value is None:
        return NA
    if isinstance(value, str) and not value.strip():
        return NA
    return value


def _round(value, digits: int = 2):
    return round(value, digits) if isinstance(value, (int, float)) else value


def build_detailed_bom(members: list[Member]) -> pd.DataFrame:
    rows = []
    for m in members:
        rows.append(
            {
                "Member_ID": _na(m.member_id),
                "Member_Type": _na(m.member_type),
                "Section_Size": _na(m.section_size),
                "Length_mm": _na(_round(m.length_mm)),
                "Quantity": _na(m.quantity),
                "Elevation_m": _na(_round(m.elevation_m)),
                "Grid_Location": _na(m.grid_location),
                "Weight_kg": _na(_round(m.weight_kg)),
                "Source_File": _na(m.source_file),
            }
        )
    return pd.DataFrame(rows, columns=DETAILED_COLUMNS)


def build_summary(members: list[Member]) -> pd.DataFrame:
    """Total quantity and total length per Section_Size."""
    buckets: dict[str, dict[str, float]] = {}
    for m in members:
        section = (m.section_size or NA)
        qty = m.quantity or 0
        bucket = buckets.setdefault(section, {"qty": 0, "length": 0.0, "weight": 0.0})
        bucket["qty"] += qty
        if isinstance(m.length_mm, (int, float)):
            bucket["length"] += m.length_mm * qty
        if isinstance(m.weight_kg, (int, float)):
            bucket["weight"] += m.weight_kg * qty

    rows = [
        {
            "Section_Size": section,
            "Total_Qty": data["qty"],
            "Total_Length_mm": round(data["length"], 2),
            "Total_Weight_kg": round(data["weight"], 2),
        }
        for section, data in sorted(buckets.items())
    ]
    return pd.DataFrame(rows, columns=["Section_Size", "Total_Qty", "Total_Length_mm", "Total_Weight_kg"])


def build_type_counts(members: list[Member]) -> pd.DataFrame:
    """Total number (pieces) of each Member_Type."""
    counts: dict[str, int] = {}
    for m in members:
        mtype = m.member_type or NA
        counts[mtype] = counts.get(mtype, 0) + (m.quantity or 0)
    rows = [{"Member_Type": t, "Total_Count": c} for t, c in sorted(counts.items())]
    return pd.DataFrame(rows, columns=["Member_Type", "Total_Count"])
