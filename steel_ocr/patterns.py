"""Regex patterns for steel member marks, sections, plates, and dimensions."""

from __future__ import annotations

import re

# Section sizes: W18x35, ISMB400, UB457x191x67, UC254x254x73, HSS6x6x3/8, etc.
SECTION_SIZE_RE = re.compile(
    r"\b(?:"
    r"W\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"
    r"|HP\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"
    r"|WT\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"
    r"|C\d{1,2}\s*[x×]\s*\d{1,2}(?:\.\d+)?"
    r"|MC\d{1,2}\s*[x×]\s*\d{1,2}(?:\.\d+)?"
    r"|L\d{1,2}\s*[x×]\s*\d{1,2}\s*[x×]\s*[\d/]+"
    r"|HSS\d{1,2}(?:\.\d+)?\s*[x×]\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d/]+"
    r"|ISMB\s?\d{2,4}"
    r"|ISMC\s?\d{2,4}"
    r"|ISWB\s?\d{2,4}"
    r"|UB\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"
    r"|UC\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"
    r"|HEA\s?\d{2,3}"
    r"|HEB\s?\d{2,3}"
    r"|IPE\s?\d{2,3}"
    r"|SHS\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,2}"
    r"|RHS\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,2}"
    r")\b",
    re.IGNORECASE,
)

# Piece marks — exclude BP (base plate) and BR/XB (bracing) from beams
BEAM_MARK_RE = re.compile(
    r"\b(?:BEAM|BM|B(?![PR]))([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)
COLUMN_MARK_RE = re.compile(
    r"\b(?:COLUMN|COL|C)([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)
BASEPLATE_MARK_RE = re.compile(
    r"\b(?:BPL|BP|BASE\s*PLATE)([-\s]?[A-Z]?\d{1,4}[A-Z]?)\b",
    re.IGNORECASE,
)

LENGTH_RE = re.compile(
    r"(?:L(?:ength)?|Ht|H(?:eight)?)\s*[:=]?\s*"
    r"("
    r"\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?"
    r"|\d{3,5}(?:\.\d+)?\s*(?:mm|m)?"
    r"|\d{1,3}(?:\.\d+)?\s*(?:ft|FT)"
    r")",
    re.IGNORECASE,
)
STANDALONE_LENGTH_RE = re.compile(
    r"\b(\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?|\d{3,5}\s*mm)\b",
    re.IGNORECASE,
)

# Explicit quantity only — never bare "x35" from W18x35 / plate 500x500x25.
QUANTITY_RE = re.compile(
    r"(?:"
    r"(?:Qty|QTY|Quantity|Count)\s*[:=]?\s*(\d{1,4})"
    r"|(\d{1,4})\s*(?:Nos?\.?|No\.?|Off|EA|Pcs?\.?)"
    r"|(?<![A-Z0-9])\(\s*(\d{1,4})\s*\)(?!\s*(?:mm|m|M\d|Ø|DIA))"
    r")",
    re.IGNORECASE,
)

MATERIAL_RE = re.compile(
    r"\b(A36|A572(?:\s*Gr\.?\s*\d+)?|A992|A500(?:\s*Gr\.?\s*[ABC])?|"
    r"S275(?:JR)?|S355(?:JR)?|ASTM\s*A\d+|IS\s*2062(?:\s*E\d+)?|"
    r"Gr\.?\s*(?:36|50|55)|FY\s*\d{2,3})\b",
    re.IGNORECASE,
)

PLATE_SIZE_RE = re.compile(
    r"(?<![A-Z])(?:PL(?:ATE)?\s*)?"
    r"(\d{2,4}(?:\.\d+)?\s*[x×]\s*\d{2,4}(?:\.\d+)?\s*[x×]\s*\d{1,3}(?:\.\d+)?)"
    r"(?:\s*mm)?",
    re.IGNORECASE,
)

ANCHOR_BOLT_RE = re.compile(
    r"(?:"
    r"(\d+)\s*[-–]\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"
    r"|(\d+)\s*(?:Nos?\.?|No\.?)\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"
    r"|\(\s*(\d+)\s*\)\s*(?:M|Ø|DIA\.?\s*)(\d{1,2}(?:\.\d+)?)"
    r"|(\d+)\s*[-–]\s*(\d/\d|\d(?:\.\d+)?)\s*[\"']?\s*(?:DIA|Ø)?"
    r")",
    re.IGNORECASE,
)

ELEVATION_RE = re.compile(
    r"(?:EL(?:EV(?:ATION)?)?|T\.?O\.?S\.?|T\.?O\.?C\.?|TOP|BASE)\s*[.:]?\s*"
    r"([+-]?\d{1,3}'\s*-?\s*\d{1,2}(?:\s*\d/\d)?\"?"
    r"|[+-]?\d{1,4}(?:\.\d{1,3})?(?:\s*(?:mm|m|ft)\b)?)",
    re.IGNORECASE,
)

# Schedule-style: B1  W12x26  4  or  Beam B3 ISMB300 Qty=2
SCHEDULE_ROW_RE = re.compile(
    r"(?P<mark>(?:BEAM|BM|B(?![PR])|COLUMN|COL|C|BPL|BP|BASE\s*PLATE)"
    r"[-\s]?[A-Z]?\d{1,4}[A-Z]?)"
    r".{0,40}?"
    r"(?P<section>"
    r"W\d{1,2}\s*[x×]\s*\d{1,3}(?:\.\d+)?"
    r"|ISMB\s?\d{2,4}|ISMC\s?\d{2,4}|ISWB\s?\d{2,4}"
    r"|UB\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"
    r"|UC\s?\d{2,3}\s*[x×]\s*\d{2,3}\s*[x×]\s*\d{1,3}"
    r"|HSS\d{1,2}(?:\.\d+)?\s*[x×]\s*\d{1,2}(?:\.\d+)?\s*[x×]\s*[\d/]+"
    r"|HEA\s?\d{2,3}|HEB\s?\d{2,3}|IPE\s?\d{2,3}"
    r"|\d{2,4}\s*[x×]\s*\d{2,4}\s*[x×]\s*\d{1,3}"
    r")?",
    re.IGNORECASE,
)
