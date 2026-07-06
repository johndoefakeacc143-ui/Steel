"""Recover member marks that are drawn as rotated / vertically-stacked text.

pdfplumber's ``extract_text`` groups characters into lines by their vertical
position, so a mark written vertically (e.g. ``BR4`` stacked character-by-
character) is scattered across several "lines" and never appears as a contiguous
token. Such marks are therefore missed by text-based parsing.

This module reconstructs marks directly from character geometry: characters are
clustered by spatial proximity, each cluster is read as a token, and known mark
patterns are matched. It returns only marks that the normal text layer did not
already contain, so accurate horizontal counts are preserved and rotated marks
(like ``BR4``) are added.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict

from src.bom_parser import DRAWING_MARK_PATTERN

_EPS = 12.0  # max centre distance (px) for two chars to join a cluster
_MAX_LABEL_CHARS = 6  # clusters larger than this are horizontal text lines
# Anchored mark pattern for validating a single reconstructed label.
_MARK_FULLMATCH = re.compile(r"(?:BR\d+|BP\d+|PB[1-9][A-Z]?|PB[A-Z]|B[2-9])$", re.I)


def _clusters(chars: list[dict]) -> list[list[dict]]:
    cells: dict[tuple[int, int], list[int]] = defaultdict(list)
    centres = []
    for idx, c in enumerate(chars):
        cx = (c["x0"] + c["x1"]) / 2.0
        cy = (c["top"] + c["bottom"]) / 2.0
        centres.append((cx, cy))
        cells[(int(cx // _EPS), int(cy // _EPS))].append(idx)

    parent = list(range(len(chars)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        parent[find(i)] = find(j)

    for (gx, gy), idxs in cells.items():
        neighbours = [
            idx
            for dx in (-1, 0, 1)
            for dy in (-1, 0, 1)
            for idx in cells.get((gx + dx, gy + dy), [])
        ]
        for i in idxs:
            for j in neighbours:
                if i < j:
                    if (
                        abs(centres[i][0] - centres[j][0]) <= _EPS
                        and abs(centres[i][1] - centres[j][1]) <= _EPS
                    ):
                        union(i, j)

    groups: dict[int, list[dict]] = defaultdict(list)
    for i in range(len(chars)):
        groups[find(i)].append(chars[i])
    return list(groups.values())


def _cluster_marks(chars: list[dict]) -> Counter:
    """Count marks from single-label clusters, trying every reading order.

    Rotated labels can read top-to-bottom or bottom-to-top; each small cluster is
    one label, so it contributes at most one mark (counted once) regardless of
    orientation.
    """
    found: Counter = Counter()
    for group in _clusters(chars):
        if len(group) > _MAX_LABEL_CHARS:
            continue
        orderings = (
            sorted(group, key=lambda c: c["top"]),
            sorted(group, key=lambda c: -c["top"]),
            sorted(group, key=lambda c: c["x0"]),
            sorted(group, key=lambda c: -c["x0"]),
        )
        for ordered in orderings:
            token = "".join(c["text"] for c in ordered).strip()
            if _MARK_FULLMATCH.fullmatch(token):
                found[token.upper()] += 1
                break
    return found


def recover_stacked_marks(chars: list[dict], existing_text: str) -> dict[str, int]:
    """Return marks found via char geometry but absent from the text layer.

    ``chars`` is ``pdfplumber`` ``page.chars``; ``existing_text`` is the page's
    ``extract_text`` output. Only marks with no occurrence in the text layer are
    returned, so horizontal counts stay authoritative.
    """
    if not chars:
        return {}
    text_marks = {m.upper() for m in DRAWING_MARK_PATTERN.findall(existing_text or "")}
    clustered = _cluster_marks(chars)
    return {mark: count for mark, count in clustered.items() if mark not in text_marks}
