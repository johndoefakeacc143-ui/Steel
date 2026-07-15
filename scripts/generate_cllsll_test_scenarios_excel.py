#!/usr/bin/env python3
"""
Generate Test Scenarios + Traceability Matrix Excel from REq_dd36.txt
(CLL/SLL BioREAD General Requirements — Sessions 1–4 + cross-session modules).

Sheets:
  - Cover
  - Test Scenarios: Requirement ID | Test Scenario ID | Test Scenario | Expected result
  - Traceability Matrix: Requirement ID | Requirement Description | Test Scenario ID | Status
  - Requirements Catalog
"""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = Path("/home/ubuntu/.cursor/projects/workspace/uploads/REq_dd36.txt")
OUT = Path(
    "/workspace/validation/test_artifacts/CLLSLL_BRS_Test_Scenarios_Traceability.xlsx"
)

REQ_START = re.compile(
    r"^(BRS-[A-Z0-9]+-\d+[a-zA-Z]?)\s+(.*)$",
    re.MULTILINE,
)

MODULE_BY_PREFIX = [
    ("BRS-DR-", "Display Requirements"),
    ("BRS-RA-", "Reader Allocation"),
    ("BRS-IDH-", "Image Data Handling"),
    ("BRS-TAA-", "Technical Adequacy"),
    ("BRS-TLL-", "Target Lesions"),
    ("BRS-NTLL-", "Non-Target Lesions"),
    ("BRS-NLL-", "New Lesions"),
    ("BRS-SSI-", "Spleen Labeling/Tracking"),
    ("BRS-LT-", "Lesion Tracking Panels"),
    ("BRS-S1DD-", "Session 1 Data Display"),
    ("BRS-S1QA-", "Session 1 Q&A"),
    ("BRS-ILR-", "Session 2 Target Response"),
    ("BRS-NIR-", "Session 2 Non-Target Response"),
    ("BRS-NEW-", "Session 2 New Lesion Impact"),
    ("BRS-SRCR-", "Session 2 Spleen Response"),
    ("BRS-SRPR-", "Session 2 Spleen Response"),
    ("BRS-SRSD-", "Session 2 Spleen Response"),
    ("BRS-SRPD-", "Session 2 Spleen Response"),
    ("BRS-SRNE-", "Session 2 Spleen Response"),
    ("BRS-SRNC-", "Session 2 Spleen Response"),
    ("BRS-SR-", "Session 2 Spleen Response"),
    ("BRS-LRCR-", "Session 2 Liver Response"),
    ("BRS-LRSD-", "Session 2 Liver Response"),
    ("BRS-LRPD-", "Session 2 Liver Response"),
    ("BRS-LRNE-", "Session 2 Liver Response"),
    ("BRS-LRNC-", "Session 2 Liver Response"),
    ("BRS-S2DD-", "Session 2 Data Display"),
    ("BRS-S2QA-", "Session 2 Q&A"),
    ("BRS-S3DD-", "Session 3 Global Display"),
    ("BRS-S3QA-", "Session 3 Global Q&A"),
    ("BRS-REA-", "Reassessment"),
    ("BRS-QAREA-", "Reassessment Q&A"),
    ("BRS-QARA-", "Reassessment Q&A"),
    ("BRS-S4DD-", "Session 4 Adjudication Display"),
    ("BRS-S4QA-", "Session 4 Adjudication Q&A"),
]


def module_for(req_id: str) -> str:
    for prefix, module in MODULE_BY_PREFIX:
        if req_id.startswith(prefix):
            return module
    return "General"


def clean_desc(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_requirements(raw: str) -> list[tuple[str, str, str]]:
    """Return ordered unique (req_id, description, module)."""
    matches = list(REQ_START.finditer(raw))
    requirements: list[tuple[str, str, str]] = []
    seen: set[str] = set()

    for i, match in enumerate(matches):
        req_id = match.group(1)
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(raw)
        block = raw[start:end]
        # strip leading ID
        body = re.sub(rf"^{re.escape(req_id)}\s+", "", block, count=1)
        # stop at next section header-ish lines that aren't content bullets
        # keep full requirement body but trim trailing whitespace
        desc = clean_desc(body)
        # collapse to a usable single-paragraph description (keep key content)
        desc_one = re.sub(r"\s+", " ", desc)
        if len(desc_one) > 1200:
            desc_one = desc_one[:1197] + "..."
        if req_id in seen:
            # Duplicate ID in source (e.g. BRS-DR-055 twice) — keep first
            continue
        # Skip dangling cross-reference only IDs with no real definition
        # (e.g. BRS-TLL-046 appears only as exception reference inside another req)
        if len(desc_one) < 20 and "exception" in desc_one.lower():
            continue
        seen.add(req_id)
        requirements.append((req_id, desc_one, module_for(req_id)))
    return requirements


def make_scenarios(req_id: str, desc: str, module: str) -> list[tuple[str, str]]:
    """Return list of (scenario, expected_result) — Positive, Negative, UI."""
    short = desc if len(desc) <= 220 else desc[:217] + "..."

    # Response-code hints for expected results
    response_hint = ""
    for code in (
        "CR",
        "PR",
        "SD",
        "PD",
        "NE",
        "NA",
        "NC",
        "NED",
        "CMR",
        "PMR",
        "NMR",
        "PMD",
        "Yes",
        "No",
    ):
        # trailing response column style: ends with tab/code
        if re.search(rf"\b{code}\b\s*$", desc) or f"\t{code}" in desc:
            response_hint = code
            break

    expected_pos = (
        f"System behaves per {req_id}; "
        + (
            f"derived/displayed result = {response_hint}. "
            if response_hint
            else "required values/options/calculations are correct. "
        )
        + "Data persists and is available for downstream panels/sessions."
    )

    scenarios = [
        (
            f"[Positive][{module}] Verify valid path for {req_id}: {short}",
            expected_pos,
        ),
        (
            f"[Negative][{module}] Attempt to violate/bypass {req_id} "
            f"(missing required input, forbidden action, out-of-sequence, or invalid combination).",
            "System blocks the action and/or shows the specified validation/error message; "
            "invalid data is not saved; sign-off prevented when required.",
        ),
        (
            f"[UI][{module}] Verify UI for {req_id}: labels, answer options, "
            f"enable/disable/hidden/read-only states, messages, and panel layout.",
            "UI controls, field states, and displayed text match the requirement; "
            "no incorrect or orphaned controls.",
        ),
    ]

    # Extra positive paths for multi-rule requirements
    extra: list[tuple[str, str]] = []
    dlow = desc.lower()

    if "auto-populate" in dlow or "auto populate" in dlow:
        extra.append(
            (
                f"[Positive][{module}] Verify auto-populate / default rules for {req_id}.",
                "Fields auto-populate and disable/enable exactly as specified.",
            )
        )
    if "sign off" in dlow or "sign-off" in dlow or "upon sign" in dlow:
        extra.append(
            (
                f"[Negative][{module}] Attempt sign-off without meeting {req_id} conditions.",
                "Sign-off blocked with the required message until corrected/acknowledged.",
            )
        )
    if "hidden" in dlow:
        extra.append(
            (
                f"[UI][{module}] Confirm hidden field(s) for {req_id} are not visible to reader but stored/exported as required.",
                "Field hidden on eCRF; value present in database/export when applicable.",
            )
        )
    if "read only" in dlow or "read-only" in dlow or "always read only" in dlow:
        extra.append(
            (
                f"[Negative][{module}] Attempt to edit read-only field governed by {req_id}.",
                "Field remains non-editable; value unchanged.",
            )
        )
    if "session 1" in dlow and "session 2" in dlow:
        extra.append(
            (
                f"[Positive][{module}] Verify Session 1→Session 2 carry-forward / sequential behavior for {req_id}.",
                "Values/labels/statuses carry forward and apply in follow-up as specified.",
            )
        )
    if "adjudicat" in dlow:
        extra.append(
            (
                f"[Positive][{module}] Verify adjudication path for {req_id} with discrepant primary readers.",
                "Only discrepant assessments are enabled for adjudicator; other fields blank/disabled as specified.",
            )
        )

    # Keep extras but avoid exploding count too much — max 2 extras
    scenarios.extend(extra[:2])
    return scenarios


def style_header(ws, color: str) -> None:
    fill = PatternFill("solid", fgColor=color)
    font = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin
    ws.row_dimensions[1].height = 22


def autosize(ws, widths: dict[int, int]) -> None:
    thin = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin


def main() -> None:
    raw = SRC.read_text(encoding="utf-8", errors="replace")
    requirements = parse_requirements(raw)

    test_rows: list[list[str]] = []
    trace_rows: list[list[str]] = []
    ts = 1
    for req_id, desc, module in requirements:
        for scenario, expected in make_scenarios(req_id, desc, module):
            ts_id = f"TS-{ts:04d}"
            ts += 1
            test_rows.append([req_id, ts_id, scenario, expected])
            trace_rows.append([req_id, desc, ts_id, "Not Executed"])

    wb = Workbook()

    # Cover
    cover = wb.active
    cover.title = "Cover"
    cover["A1"] = (
        "CLL/SLL BioREAD Windows Application — Test Scenarios & Traceability Matrix"
    )
    cover["A1"].font = Font(bold=True, size=14, color="0B6E6E")
    rows = [
        ("Source document", "REq_dd36.txt — General Requirements (CLL/SLL)"),
        ("Disease indication", "CLL/SLL"),
        (
            "Sessions covered",
            "Session 1 Screening, Session 2 Follow-up, Session 3 Global, Session 4 Adjudication "
            "(+ Display, Reader Allocation, Image Handling, TAA, Lesions, Spleen, Liver, Reassessment)",
        ),
        ("Case types", "Positive, Negative, UI"),
        ("Test Scenario ID format", "TS-0001"),
        ("Total requirements", str(len(requirements))),
        ("Total test scenarios", str(len(test_rows))),
        ("Traceability status default", "Not Executed"),
        (
            "Sheets",
            "Cover | Test Scenarios | Traceability Matrix | Requirements Catalog",
        ),
    ]
    r = 3
    for label, value in rows:
        cover[f"A{r}"] = label
        cover[f"B{r}"] = value
        r += 1
    cover.column_dimensions["A"].width = 30
    cover.column_dimensions["B"].width = 110

    # Test Scenarios
    ws1 = wb.create_sheet("Test Scenarios")
    ws1.append(
        ["Requirement ID", "Test Scenario ID", "Test Scenario", "Expected result"]
    )
    for row in test_rows:
        ws1.append(row)
    style_header(ws1, "0B6E6E")
    autosize(ws1, {1: 18, 2: 16, 3: 90, 4: 55})
    ws1.auto_filter.ref = ws1.dimensions
    ws1.freeze_panes = "A2"

    # Traceability Matrix
    ws2 = wb.create_sheet("Traceability Matrix")
    ws2.append(
        [
            "Requirement ID",
            "Requirement Description",
            "Test Scenario ID",
            "Status",
        ]
    )
    for row in trace_rows:
        ws2.append(row)
    style_header(ws2, "0A4F6E")
    autosize(ws2, {1: 18, 2: 95, 3: 16, 4: 16})
    ws2.auto_filter.ref = ws2.dimensions
    ws2.freeze_panes = "A2"

    # Requirements Catalog
    ws3 = wb.create_sheet("Requirements Catalog")
    ws3.append(["Requirement ID", "Requirement Description", "Module / Session"])
    for req_id, desc, module in requirements:
        ws3.append([req_id, desc, module])
    style_header(ws3, "334155")
    autosize(ws3, {1: 18, 2: 100, 3: 34})
    ws3.auto_filter.ref = ws3.dimensions
    ws3.freeze_panes = "A2"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)

    # Also copy source + artifacts
    src_copy = OUT.parent / "General_Requirements_CLLSLL_REq.txt"
    src_copy.write_text(raw, encoding="utf-8")

    artifact_dir = Path("/opt/cursor/artifacts")
    artifact_dir.mkdir(parents=True, exist_ok=True)
    import shutil

    shutil.copy(OUT, artifact_dir / OUT.name)
    shutil.copy(src_copy, artifact_dir / src_copy.name)

    print(f"Wrote {OUT}")
    print(f"Requirements: {len(requirements)}")
    print(f"Test scenarios: {len(test_rows)}")
    # case type counts
    from collections import Counter

    c = Counter()
    for _, _, scenario, _ in ((r[0], r[1], r[2], r[3]) for r in test_rows):
        if "[Positive]" in scenario:
            c["Positive"] += 1
        elif "[Negative]" in scenario:
            c["Negative"] += 1
        elif "[UI]" in scenario:
            c["UI"] += 1
    print("Case types:", dict(c))
    print("First 5 reqs:", [r[0] for r in requirements[:5]])
    print("Last 5 reqs:", [r[0] for r in requirements[-5:]])


if __name__ == "__main__":
    main()
