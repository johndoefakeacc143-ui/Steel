#!/usr/bin/env python3
"""
Generate Lugano Criteria Windows Application test scenarios Excel workbook
from BR-S requirements document (Sessions 1–5).

Sheets:
  1. Test Scenarios  — Requirement ID | Test Scenario ID | Test Scenario | Expected result
  2. Traceability Matrix — Requirement ID | Requirement Description | Test Scenario ID | Status
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Requirements copied from General Requirements (BR-S-DR) document
# ---------------------------------------------------------------------------

REQUIREMENTS: list[tuple[str, str, str]] = [
    # (Requirement ID, Description, Session/Module)
    (
        "BRS-DR-005",
        "Display the Protocol number [R1979-ONC-2105] in BioREAD from the Read Project Name field stored in the ReadProject table within the project selection screen.",
        "General / Project Selection",
    ),
    (
        "BRS-DR-041",
        "The system must display the data via the BiopsyReport form (uploaded at the baseline/screening time point) in BioPACS if available in Session 3: Global and Session 4: Adjudication.",
        "General / BioPACS",
    ),
    (
        "BRS-RA-005",
        "A case is selected for adjudication if discrepant between the two primary readers when answers do not match exactly for FinalTimePointResponsePET, FinalTimePointResponseCT and/or FinalTimePointResponseCTPET at any and all time points.",
        "Reader Allocation",
    ),
    (
        "BRS-TAA-005",
        "Sessions 1 & 2: Ask reason for images Readable but not Optimal, or Not Readable (ImageAdequacyReason). One answer required, multiple allowed when enabled. Auto-populate/disable options based on image availability (Time point missing, Missing PET Whole Body, Other specify).",
        "Technical Adequacy (S1/S2)",
    ),
    (
        "BRS-LLT-005",
        "When Extranodal site location Muscle/Soft Tissue/Subcutis is selected (Index/Non-Index/New/PET P1), require response for Muscle/Soft Tissue/Subcutis specify with defined anatomical options.",
        "Lesion Labeling/Tracking",
    ),
    (
        "BRS-ILL-015",
        "Session 1 Target lesion context menu: Previously Irradiated? (No/Yes/Yes with progression). Default No. If Yes, block Target add and show message to reassign as non-target. Carry over to Session 2.",
        "Session 1 Target Lesions",
    ),
    (
        "BRS-ILL-025",
        "Session 2 eCRF context menu to assign Index/Target as missing or resolved. Do not provide Not present, resolved if no CT or MR images available.",
        "Session 2 Target Lesions",
    ),
    (
        "BRS-ILL-055",
        "Index/Target with status Merged-Not Present [label] must not count as missing when determining time point response. LDi/SDi defaulted to 0 (not a calculation).",
        "Session 2 Target Lesions",
    ),
    (
        "BRS-NIL-015",
        "Session 1 Non-Target context menu: Previously Irradiated? (No/Yes/Yes with progression). Default No. Response carried over to Session 2.",
        "Session 1 Non-Target Lesions",
    ),
    (
        "BRS-NIL-025",
        "Session 2 eCRF context menu to assign Non-Index/Non-Target as missing or resolved. Do not provide Not present, resolved if no CT or MR images available.",
        "Session 2 Non-Target Lesions",
    ),
    (
        "BRS-NLL-020",
        "Session 2: Context menu to label ROI from RT struct in MIM as New Lesion (nodal/extranodal) when threshold measurements are met in follow-up time points.",
        "Session 2 New Lesions",
    ),
    (
        "BRS-NLL-021",
        "Session 2: Context menu to label ROI as New Lesion (extranodal). If LDi < 10 mm, display advisory message recommending wait until >10 mm; reader clicks OK; no change required.",
        "Session 2 New Lesions",
    ),
    (
        "BRS-SSI-005",
        "Reader captures spleen first/last slice via MIM 2D Brush ROIs named SpleenFS/SpleenLS; eCRF displays slice numbers and calculates LVD.",
        "Spleen Scan Interval",
    ),
    (
        "BRS-SSI-035",
        "Store to hidden fields: SpleenFirstSliceLocation, SpleenLastSliceLocation, SpleenNextToLastSliceLocation, SpleenSliceThickness (blank if LVD = n/a).",
        "Spleen Scan Interval",
    ),
    (
        "BRS-LC-001",
        "Display Target Lesion Calculation Summary for CT/MR in Sessions 2 and 3 with time point, label, location, slice, Actual/Final LDi/SDi, Final PPD, Absolute Change Nadir LDi/SDi, % Change Screening/Nadir PPD, Lesion Status, PET Assessment.",
        "Tracking Panels (S2/S3)",
    ),
    (
        "BRS-LC-040",
        "Informational panel Muscle/Soft Tissue/Subcutis Specify in Sessions 2, 3, 4 and 5 when site location Muscle/Soft Tissue/Subcutis is added; else hide. Shows Label, Location, Specify.",
        "Tracking Panels (S2–S5)",
    ),
    (
        "BRS-TPS-001",
        "Timepoint Summary Panel for CT evaluation in Sessions 2 and 3 including Target/Non-Target Response, new lesions, Spleen, CT TPR, PET+CT TPR, Designated PET+CT TPR, comments, reassessment reason.",
        "Time Point Summary (S2/S3)",
    ),
    (
        "BRS-TPS-005",
        "Timepoint Summary Panel for PET evaluation in Sessions 2 and 3 including bone marrow, Five-Point Scale, qualitative change, PET TPR, Designated PET TPR, comments.",
        "Time Point Summary (S2/S3)",
    ),
    (
        "BRS-S1QA-000a",
        "Session 1: Indicate subjects FDG-avidity (FDG-Avid / Non FDG-Avid / NE). Required when Optimal or Readable but not optimal. Default NE when Not Readable or no PET. Block ROI add until answered. FDG-Avid requires PET Positive lesion (P1) at sign-off.",
        "Session 1 Q&A",
    ),
    (
        "BRS-S1QA-002",
        "Session 1: Evidence of bone marrow involvement. Required when Optimal/Readable. Auto NE when Not Readable or no PET. Block Non FDG-avid + positive marrow at sign-off.",
        "Session 1 Q&A",
    ),
    (
        "BRS-LM-035",
        "Session 2 Index/Target lesions decreasing to <5 mm in any dimension categorized as TSTM; assign 5 mm for that axis (except when BR-S-LM-036 Yes). Also applies to split fragments below threshold.",
        "Session 2 Lesion Measurements",
    ),
    (
        "BRS-LM-036",
        "Session 2: When LDi and/or SDi is >0 and <5 mm, ask Was the lesion reliably measured? Yes=use true values; No=apply TSTM (BRS-LM-035).",
        "Session 2 Lesion Measurements",
    ),
    (
        "BRS-CICR-001",
        "CT Index CR: All extranodal Index/Target Not Present/Resolved and all nodal Index LDi ≤ 15 mm.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CIPR-001",
        "CT Index PR: % SPD Change from Screening decrease ≥50% and criteria for CR, PD or NE not met.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CISD-001",
        "CT Index SD: Populate SD if CR, PR, PD, NE or NA criteria have not been met.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CIPD-001",
        "CT Index PD is derived based on a single lesion instance, not on lesion sums.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CIPD-005",
        "CT Index PD requires: (1) abnormal LDi >15 mm; (2) ≥50% PPD increase from nadir; (3) absolute LDi/SDi increase thresholds from nadir (5 mm if nadir ≤20 mm, 10 mm if >20 mm). Special rules for split/merged lesions.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CIPD-010",
        "CT Index PD: Any Extranodal lesion resolves then at subsequent TP re-assigned Present and measures >10 mm LDi.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CINE-001",
        "CT Index NE when time point marked Not Readable and Index/Target lesions were present at screening/baseline.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CINE-005",
        "CT Index NE if PPD of at least one Index/Target lesion is not evaluable and PD criteria not met.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-CINA-001",
        "CT Index NA when no Index/Target disease at screening/baseline and current TP is not Not Readable.",
        "Session 2 CT Index Response",
    ),
    (
        "BRS-SR-000",
        "Spleen Nadir is smallest LVD including screening up to but not including current time point.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SR-001",
        "Spleen Response based on first time point a spleen measurement is assessable as Reference.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SR-002",
        "Spleen LVD >130 mm = abnormal; LVD ≤130 mm = normal.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SR-003",
        "Reader can edit Enlarged Spleen Status to Normal; Normal designation takes precedence over measurement designation.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRCR-001",
        "Spleen CR: Abnormal spleen (ASV >0) at Reference regressed to normal (LVD ≤130 mm).",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRCR-002",
        "Spleen CR: Reader changes Spleen Status from Enlarged to Normal when abnormal at nadir prior TP.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRPR-001",
        "Spleen PR: Abnormal at Reference remains abnormal but ASV regressed >50% from Reference.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRSD-001",
        "Spleen SD: Abnormal at Reference remains abnormal; neither PR nor PD met.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRPD-020",
        "Spleen PD: Normal at nadir becomes abnormal (LVD >130) and increased ≥20 mm LVD vs nadir.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRPD-025",
        "Spleen PD: Abnormal at nadir increased >50% ASV vs nadir and ≥10 mm LVD.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNE-001",
        "Spleen NE when time point marked Not Readable.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNE-005",
        "Spleen NE when Not Present due to technical issues or Not Present – images not available.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNE-010",
        "Spleen NE when prior measurement exists and current status is Not Present, Surgery.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNE-015",
        "Spleen NE if unevaluable at screening and abnormal (>130 mm) at first on-study measurement.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNC-001",
        "Spleen NC: Normal at Reference remains normal.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNC-005",
        "Spleen NC: Screening unevaluable but spleen normal (≤130) at first on-study assessable TP, or status changed Enlarged to Normal.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNC-010",
        "Spleen NC: Normal at Reference becomes abnormal but does not meet PD.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNC-015",
        "Spleen NC: Status Not Present, surgery at screening.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-SRNC-020",
        "Spleen NC: Screening unevaluable and Not Present, surgery at first assessable on-study TP.",
        "Session 2 Spleen Response",
    ),
    (
        "BRS-S2DD-025",
        "Session 2: Display Most FDG-avid Lesion (P1) info: Lesion, SUVmax, % Change SUVmax Screening, % Change SUVmax Nadir. Auto NE when PET lesion identified but no SUV values.",
        "Session 2 Data Display",
    ),
    (
        "BRS-S2QA-002",
        "Session 2 bone marrow involvement question with additional New or recurrent option; real-time and sign-off validation messages vs prior TP and P1/PET TPR.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-010b",
        "Session 2: Qualitative combined changes in overall intensity (Increase from nadir / Decrease from screening / No Change / NE) with auto-populate rules based on Five Point Score, screening PET status, and new CT PET-positive lesions.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-020",
        "Session 2: PET Time Point Response (read-only CMR/PMR/NMR/PMD/NA/NE) with auto-populate rules including prior PMD lock, bone marrow new/recurrent → PMD, out-of-sequence NE rules.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-022",
        "Session 2: Designated PET TPR — defaults to PET TPR; edit limited by FDG avidity, prior designated responses, and disable rules for NE/PMD.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-021",
        "Session 2: Hidden PET TPR Date — latest date if Designated PET ≠ PMD; earliest if = PMD; n/a if Not Readable due to Time Point missing.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-023",
        "Session 2: Comment on changing PET TPR required when PET TPR and Designated PET TPR do not match.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-045",
        "Session 2: CT Time Point Response (read-only) combined from Target/Non-Target/New lesions/Spleen with out-of-sequence NE and reader-choice rows.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-050",
        "Session 2: PET+CT Time Point Response (read-only) combined from CT TPR + PET TPR per BR-S-IRTPR-010; NE if Not Readable or PET out-of-sequence NE.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-052",
        "Session 2: Designated PET+CT TPR — editable with avidity-specific options, prior CR/PR option limits, disable for NE/PD.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-051",
        "Session 2: Hidden PET+CT TPR Date — latest if Designated ≠ PD; earliest if = PD; n/a if Time Point missing.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S2QA-053",
        "Session 2: Comment on changing PET+CT TPR required when PET+CT TPR and Designated PET+CT TPR do not match.",
        "Session 2 Q&A",
    ),
    (
        "BRS-S3LC-001",
        "Session 3: Target Lesion Calculation Summary with interaction on Lesion Status plus Override Status/Reason fields.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3LC-005",
        "Session 3: Target Lesions Sums — Response SPD, % Change Screening SPD, % Change Nadir SPD with n/a rules.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3LC-020",
        "Session 3: Non-Target lesion panel with Status, PET Assessment, Override Status/Reason and interaction.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3TPS-001",
        "Session 3 CT Timepoint Summary Panel including Override CT Target/Non-Target Response and Final CT TPR.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3TPS-005",
        "Session 3 PET Timepoint Summary Panel with interaction on Designated PET TPR, Override Designated PET TPR, Final PET TPR.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3TPS-010",
        "Session 3 TPR Summary Panel for PET+CT with CT/PET/PET+CT designated and override/final fields.",
        "Session 3 Global Display",
    ),
    (
        "BRS-S3QA-010",
        "Session 3: Is there on-study radiotherapy and/or surgical data present? (Yes/No) required; maintain answer on re-assess/insert.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-011",
        "Session 3: Does on-study radiotherapy data impact CT assessment? Enabled when S3QA-010=Yes; sign-off checks vs lesion overrides.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-012",
        "Session 3: Does on-study RT/surgery impact PET assessment? Enabled when S3QA-010=Yes; Yes/No/Yes with no change required; sign-off checks vs PET overrides.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-013",
        "Session 3: Always display instructional field text about performing lesion and PET TPR override when on-study clinical data impacts TPR.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-014",
        "Session 3: Override Lesion Status on Target panel when RT present and impacts CT; carry forward; undo rules with required comment.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-015",
        "Session 3: Override Target Response = NE when any target override NE radiotherapy and Target Response not PD; PD remains.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-016",
        "Session 3: Final Target Response = Override if present else actual IndexResponseCT.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-017",
        "Session 3: Override Lesion Status on Non-Target panel (Present/Present and Normal only); carry forward; undo comment rules.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-018",
        "Session 3: Override Non-Target Response = NE when override present and Non-Target Response not PD.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-019",
        "Session 3: Final Non-Target Response = Override if present else actual NonIndexResponseCT. New lesions have no override.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-020",
        "Session 3: Override CT TPR = NE when Target/Non-Target override Yes and CT TPR not PD; PD not overridden.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-021",
        "Session 3: Override CT TPR Date — latest exam date when overridden and TPR ≠ PD.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-022",
        "Session 3: Final CT TPR = Override CT TPR if present else actual TimePointResponseCT.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-023",
        "Session 3: Final CT TPR Date = override date if override used else actual CT TPR date.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-024",
        "Session 3: Override Designated PET TPR when RT/surgery impacts PET and Designated PET ≠ PMD; carry forward; undo comment rules.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-025",
        "Session 3: Override PET TPR Date — latest date when Designated PET overridden and ≠ PMD.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-026",
        "Session 3: Final PET TPR = Override Designated PET if present else Designated PET TPR.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-027",
        "Session 3: Final PET TPR Date = override date if override used else Designated PET date.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-028",
        "Session 3: Override Designated PET+CT TPR = NE when CT and/or PET override Yes and current Designated PET+CT ≠ PD.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-029",
        "Session 3: Override Designated PET+CT TPR Date — latest when overridden and ≠ PMD/PD rules.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-030",
        "Session 3: Final PET+CT TPR = Override if present else Designated PET+CT TPR.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-030a",
        "Session 3: Final PET+CT TPR Date = override date if override used else Designated PET+CT date.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-031",
        "Session 3: Visible required comment when undoing lesion status and/or PET TPR override Yes→No from prior Global session.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3QA-032",
        "Session 3: Hidden field stores concatenated override change comment history pipe-separated per Global session.",
        "Session 3 Global Q&A",
    ),
    (
        "BRS-S3RI-001",
        "Session 3: PET Date of Progression = first Final PET TPR date where PMD achieved, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-005",
        "Session 3: CT Date of Progression = first Final CT TPR date where PD achieved, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-010",
        "Session 3: PET Date of First Response = first Final PET of PMR or CMR (whichever first) when Best is CMR/PMR, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-015",
        "Session 3: CT Date of First Response = first Final CT of PR or CR when Best is CR/PR, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-020",
        "Session 3: PET+CT Date of Progression = first Final PET+CT TPR date where PD achieved, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-025",
        "Session 3: PET+CT Date of First Response = first Final PET+CT of PR or CR when Best is CR/PR, else n/a.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-030",
        "Session 3: PET Best Response = best Final PET TPR prior to or including PMD (CMR>PMR>NMR>PMD>NA>NE). Baseline-only = NE.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-035",
        "Session 3: CT Best Response = best Final CT TPR prior to or including PD (NED>NA>CR>PR>SD>PD>NE). Baseline-only = NE.",
        "Session 3 Response Information",
    ),
    (
        "BRS-S3RI-040",
        "Session 3: PET+CT Best Response = best Final PET+CT prior to or including PD (NA/NED>CR>PR>SD>PD>NE). Baseline-only = NE.",
        "Session 3 Response Information",
    ),
    (
        "BRS-REA-001",
        "Reassessment: Allow reassessment of follow-up TP; Session 1 FDG-avidity and Technical Adequacy locked; Baseline READ ONLY once completed.",
        "Reassessment",
    ),
    (
        "BRS-QARA-010",
        "Reassessment: Is this Pseudoprogression? Visible only if TPR is PD or PMD. Yes changes PD/PMD to NE and recalculates; No with once-PD-always-PD recalculates underlying response.",
        "Reassessment Q&A",
    ),
    (
        "BRS-S4DD-001",
        "Session 4 Adjudication: Display CT Time Point Summary per primary reader including overrides, Final CT TPR, Pseudoprogression.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-S4DD-005",
        "Session 4: Display PET Time Point Summary per primary reader including Final PET TPR and screening marrow/5PS.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-S4DD-006",
        "Session 4: Display TPR Summary Panel for PET+CT per primary reader with overrides and Final PET+CT TPR.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-S4DD-020",
        "Session 4: Display each primary reader CT Target lesion measurements and override fields per time point.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-S4DD-030",
        "Session 4: Display each primary reader CT Non-Target lesions with status, PET assessment, overrides.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-S4DD-050",
        "Session 4: Display PET/CT/PET+CT Best Response, Dates of First Response/Progression and Global review comment per primary reader.",
        "Session 4 Adjudication",
    ),
    (
        "BRS-OS5DD-002",
        "Session 5 Clinical: READ ONLY text listing Global session overturned timepoints, or None.",
        "Session 5 Clinical Display",
    ),
    (
        "BRS-OS5DD-025",
        "Session 5: Radiology Time Point Summary with short TPR answers, marrow, designated responses, comments, Pseudoprogression, and Most FDG-avid Lesion Summary.",
        "Session 5 Clinical Display",
    ),
    (
        "BRS-OS5DD-050",
        "Session 5: Chronological time point rows with Overall PET/CT/PET+CT+Clinical responses, Designated, comments, Clinical Data/Exam Date.",
        "Session 5 Clinical Display",
    ),
    (
        "BRS-OS5DD-061",
        "Session 5: Informational panel Reference Notes (Appendix A); toggle on/off.",
        "Session 5 Clinical Display",
    ),
    (
        "BRS-OS5DD-062",
        "Session 5: Informational panel PET + CT + Clinical Assessment (Appendix B); toggle on/off.",
        "Session 5 Clinical Display",
    ),
    (
        "BRS-OS5QA-025",
        "Session 5: Overall CT + Clinical TPR editable when clinical change needed = Yes; recalculates CT clinical progression/first response/best.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-035",
        "Session 5: Overall PET+CT+Clinical TPR read-only; recalculates related dates/best when change occurs. NED option only if P1 Negative at screening.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-036",
        "Session 5: Designated Overall PET+CT+Clinical TPR editable when clinical change needed = Yes; defaults to Overall; recalculates dates/best.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-S2QA-037",
        "Session 5: Comment on changing Overall PET+CT+Clinical TPR required when Overall and Designated do not match.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-040",
        "Session 5: Clinical Data multi-select required on rows where Response and/or Exam Date changed.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-045",
        "Session 5: Clinical Exam Date enabled when clinical change needed = Yes; date window rules by time point; default radiology max exam date.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-065",
        "Session 5: Hidden Overall PET+CT+Clinical TPR Date rules for CR/PR/SD/NED/NE vs PD (radiology vs clinical dates).",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-070",
        "Session 5: If clinical assessment needed after last radiology TP = Yes, add clinical-only time point with required responses/dates/data.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-090",
        "Session 5: Overall PET+CT+Clinical Date of Progression = first Designated Overall PD date, else n/a.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-105",
        "Session 5: Overall PET+CT+Clinical Date of First Response = first Designated CR or PR when Best is CR/PR, else n/a.",
        "Session 5 Clinical Q&A",
    ),
    (
        "BRS-OS5QA-120",
        "Session 5: Overall CT+PET+Clinical Best Response hierarchy NA/NED>CR>PR>SD>PD>NE prior to or including PD.",
        "Session 5 Clinical Q&A",
    ),
]


def scenarios_for(req_id: str, desc: str, module: str) -> list[tuple[str, str, str]]:
    """
    Build positive / negative / UI scenarios for a requirement.
    Returns list of (scenario_text, expected_result, case_type).
    """
    scenarios: list[tuple[str, str, str]] = []

    # --- Generic coverage then specialized overrides ---
    specialized = SPECIALIZED.get(req_id)
    if specialized:
        return specialized

    # Default triad when no specialized set exists
    scenarios.append(
        (
            f"[Positive][{module}] Verify system correctly implements: {desc[:180]}",
            "System behaves per requirement; expected values/options/calculations display and persist correctly.",
            "Positive",
        )
    )
    scenarios.append(
        (
            f"[Negative][{module}] Attempt to violate or bypass rule for {req_id} (invalid/missing input, out-of-sequence, or forbidden action).",
            "System blocks action and/or shows the required validation/error message; invalid data is not saved.",
            "Negative",
        )
    )
    scenarios.append(
        (
            f"[UI][{module}] Verify UI controls, labels, enable/disable/hidden states, and layout for {req_id} match specification.",
            "UI elements, field states, and displayed text match the requirement; no orphaned or incorrect controls.",
            "UI",
        )
    )
    return scenarios


SPECIALIZED: dict[str, list[tuple[str, str, str]]] = {}


def _add(req: str, items: list[tuple[str, str, str]]) -> None:
    SPECIALIZED[req] = items


# Populate specialized scenarios for high-value / multi-path requirements
_add(
    "BRS-DR-005",
    [
        (
            "[Positive][Project Selection] Open project selection screen for protocol R1979-ONC-2105 and verify Protocol number displays from Read Project Name (ReadProject table).",
            "Protocol number R1979-ONC-2105 is displayed correctly in BioREAD project selection.",
            "Positive",
        ),
        (
            "[Negative][Project Selection] Open a project where Read Project Name / ReadProject mapping is missing or mismatched.",
            "Protocol number is not incorrectly shown; system handles missing mapping without crash (blank/error per design).",
            "Negative",
        ),
        (
            "[UI][Project Selection] Verify Protocol number label/control placement and readability on project selection screen.",
            "Protocol number field is visible, correctly labeled, and formatted on the project selection UI.",
            "UI",
        ),
    ],
)

_add(
    "BRS-DR-041",
    [
        (
            "[Positive][Session 3 Global] With BiopsyReport uploaded at baseline/screening, open Session 3 and verify BiopsyReport form data displays in BioPACS.",
            "BiopsyReport data is available/displayed in BioPACS during Session 3 Global.",
            "Positive",
        ),
        (
            "[Positive][Session 4 Adjudication] With BiopsyReport uploaded at baseline, open Session 4 and verify BiopsyReport form data displays in BioPACS.",
            "BiopsyReport data is available/displayed in BioPACS during Session 4 Adjudication.",
            "Positive",
        ),
        (
            "[Negative][Session 3/4] Open Session 3/4 when no BiopsyReport was uploaded at baseline/screening.",
            "System does not display BiopsyReport content; no error crash; form remains unavailable as expected.",
            "Negative",
        ),
        (
            "[UI][Session 3/4] Verify BiopsyReport/BioPACS display controls appear only when report is available.",
            "UI shows BiopsyReport access when available and hides/disables appropriately when unavailable.",
            "UI",
        ),
    ],
)

_add(
    "BRS-RA-005",
    [
        (
            "[Positive][Reader Allocation] Two primary readers disagree on FinalTimePointResponsePET at any TP → case routed to adjudication.",
            "Case is selected for adjudication due to PET TPR discrepancy.",
            "Positive",
        ),
        (
            "[Positive][Reader Allocation] Readers disagree on FinalTimePointResponseCT → adjudication.",
            "Case is selected for adjudication due to CT TPR discrepancy.",
            "Positive",
        ),
        (
            "[Positive][Reader Allocation] Readers disagree on FinalTimePointResponseCTPET → adjudication.",
            "Case is selected for adjudication due to PET+CT TPR discrepancy.",
            "Positive",
        ),
        (
            "[Negative][Reader Allocation] Both primary readers match exactly on Final PET, CT, and PET+CT TPRs at all time points.",
            "Case is NOT selected for adjudication.",
            "Negative",
        ),
        (
            "[UI][Reader Allocation] Verify adjudication queue/status indicator when discrepancy exists vs when not.",
            "UI correctly reflects adjudication eligibility status for the case.",
            "UI",
        ),
    ],
)

_add(
    "BRS-TAA-005",
    [
        (
            "[Positive][S1/S2 Technical Adequacy] Select Readable but not Optimal or Not Readable and answer ImageAdequacyReason with one or more valid options (e.g., Artifact).",
            "Reason is accepted; multiple selections allowed when enabled; session proceeds.",
            "Positive",
        ),
        (
            "[Positive][S1/S2 Technical Adequacy] No image files for time point → Time point missing auto-populated.",
            "ImageAdequacyReason auto-populates Time point missing.",
            "Positive",
        ),
        (
            "[Positive][S1/S2 Technical Adequacy] At least one image available → Time point missing disabled.",
            "Time point missing option is disabled.",
            "Positive",
        ),
        (
            "[Positive][S1/S2 Technical Adequacy] PET images loaded → Missing PET Whole Body hidden; no PET → Missing PET Whole Body auto-populated and additional options allowed; PET/Background status Not Present, Images not Available.",
            "Missing PET Whole Body visibility/auto-populate and PET status behave per logic.",
            "Positive",
        ),
        (
            "[Positive][S1/S2 Technical Adequacy] Select Other, specify → OtherSpecify text box required and completed.",
            "OtherSpecify is enabled/required and accepts text; sign-off allowed after entry.",
            "Positive",
        ),
        (
            "[Negative][S1/S2 Technical Adequacy] Leave ImageAdequacyReason unanswered when required and attempt sign-off/continue.",
            "System blocks and requires at least one reason.",
            "Negative",
        ),
        (
            "[Negative][S1/S2 Technical Adequacy] Select Other, specify without entering OtherSpecify text and attempt continue.",
            "System blocks until OtherSpecify is completed.",
            "Negative",
        ),
        (
            "[UI][S1/S2 Technical Adequacy] Verify all listed ImageAdequacyReason options render and enable/disable/hide per image availability.",
            "All answer options and OtherSpecify control states match specification visually.",
            "UI",
        ),
    ],
)

_add(
    "BRS-LLT-005",
    [
        (
            "[Positive][Lesion Labeling] Add Index/Target extranodal lesion with location Muscle and select a specify option (e.g., Thigh Left).",
            "Muscle/Soft Tissue/Subcutis specify is required and saved with selected anatomical option.",
            "Positive",
        ),
        (
            "[Positive][Lesion Labeling] Add Non-Index and New/PET(P1) Soft Tissue/Subcutis lesions and complete specify list.",
            "Specify response required and saved for Non-Index and New/PET lesions.",
            "Positive",
        ),
        (
            "[Negative][Lesion Labeling] Select Muscle/Soft Tissue/Subcutis without selecting a specify option and attempt save.",
            "System requires specify response and blocks save.",
            "Negative",
        ),
        (
            "[UI][Lesion Labeling] Verify full specify option list displays (abdominal quadrants, limbs, face, back, etc.) and control is hidden for other extranodal sites.",
            "Option list matches requirement; control hidden when location is not Muscle/Soft Tissue/Subcutis.",
            "UI",
        ),
    ],
)

_add(
    "BRS-ILL-015",
    [
        (
            "[Positive][Session 1 Target] Add Target lesion with PreviouslyIrradiated = No (default).",
            "Lesion is added successfully; default is No.",
            "Positive",
        ),
        (
            "[Negative][Session 1 Target] Set PreviouslyIrradiated = Yes and attempt to add as Target.",
            "Lesion not added; message: Previously irradiated lesion can only be selected as a non-target lesion. Please reassign the lesion as a non-target.",
            "Negative",
        ),
        (
            "[Positive][Session 1 Target] Set PreviouslyIrradiated = Yes, with progression and verify Target add blocked same as Yes.",
            "Target add blocked with required message; change required.",
            "Positive",
        ),
        (
            "[UI][Session 1→2] Verify PreviouslyIrradiated context menu options and carry-over display in Session 2 sequential TPs.",
            "Control shows No/Yes/Yes with progression; value carried to Session 2 time points.",
            "UI",
        ),
    ],
)

_add(
    "BRS-ILL-025",
    [
        (
            "[Positive][Session 2 Target] With CT/MR available, open context menu and assign Index/Target as missing or resolved / Not present, resolved.",
            "Missing/resolved options available and status updates correctly.",
            "Positive",
        ),
        (
            "[Negative][Session 2 Target] No CT or MR images for TP → Not present, resolved option not provided.",
            "Not present, resolved is unavailable when CT/MR images are absent.",
            "Negative",
        ),
        (
            "[UI][Session 2 Target] Verify context menu options for missing/resolved based on image availability.",
            "Menu options enable/disable per CT/MR availability.",
            "UI",
        ),
    ],
)

_add(
    "BRS-ILL-055",
    [
        (
            "[Positive][Session 2 Target] Lesion status Merged-Not Present [label]: verify it does not count as missing for TPR; LDi/SDi default to 0.",
            "Merged-Not Present does not count as missing; LDi/SDi show 0 (not calculated).",
            "Positive",
        ),
        (
            "[Negative][Session 2 Target] Confirm TPR is not incorrectly forced to NE solely due to Merged-Not Present status.",
            "TPR determination ignores Merged-Not Present as missing.",
            "Negative",
        ),
        (
            "[UI][Session 2 Target] Verify Merged-Not Present label and 0 mm LDi/SDi display on tracking panel.",
            "Status label and zero diameters display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-NIL-015",
    [
        (
            "[Positive][Session 1 Non-Target] Add Non-Target with PreviouslyIrradiated default No; edit to Yes / Yes with progression.",
            "Non-Target added; PreviouslyIrradiated saved and editable; defaults to No.",
            "Positive",
        ),
        (
            "[UI][Session 1→2] Verify PreviouslyIrradiated carried over to Session 2 sequential time points for Non-Target.",
            "Value carried over and displayed in Session 2.",
            "UI",
        ),
        (
            "[Negative][Session 1 Non-Target] Leave PreviouslyIrradiated unanswered if system requires selection (clear default) and attempt save.",
            "System enforces required selection / restores default per design; invalid blank not retained.",
            "Negative",
        ),
    ],
)

_add(
    "BRS-NIL-025",
    [
        (
            "[Positive][Session 2 Non-Target] With CT/MR available, assign Non-Target missing or resolved including Not present, resolved.",
            "Status options available and applied.",
            "Positive",
        ),
        (
            "[Negative][Session 2 Non-Target] No CT/MR images → Not present, resolved not provided.",
            "Not present, resolved unavailable without CT/MR.",
            "Negative",
        ),
        (
            "[UI][Session 2 Non-Target] Verify context menu option set based on image availability.",
            "UI menu matches requirement.",
            "UI",
        ),
    ],
)

_add(
    "BRS-NLL-020",
    [
        (
            "[Positive][Session 2 New Lesion] In follow-up TP, label MIM RT-struct ROI as New Lesion (nodal/extranodal) when threshold met.",
            "ROI labeled as New Lesion successfully when thresholds met.",
            "Positive",
        ),
        (
            "[Negative][Session 2 New Lesion] Attempt New Lesion label when threshold measurements not met.",
            "System prevents New Lesion labeling until thresholds are met.",
            "Negative",
        ),
        (
            "[UI][Session 2 New Lesion] Verify New Lesion (nodal/extranodal) context menu options on eligible ROI.",
            "Context menu displays New Lesion nodal/extranodal options appropriately.",
            "UI",
        ),
    ],
)

_add(
    "BRS-NLL-021",
    [
        (
            "[Positive][Session 2 New Lesion] Add new extranodal lesion with LDi < 10 mm; acknowledge OK message; lesion remains.",
            "Message displayed: New extranodal lesions may be of any size... recommended to wait until >10 mm. OK closes message; no change required; lesion retained.",
            "Positive",
        ),
        (
            "[Positive][Session 2 New Lesion] Add new extranodal lesion with LDi > 10 mm; no advisory required.",
            "Lesion added without <10 mm advisory message.",
            "Positive",
        ),
        (
            "[UI][Session 2 New Lesion] Verify advisory dialog text and OK button behavior for LDi < 10 mm extranodal new lesion.",
            "Dialog text matches requirement; OK dismisses dialog.",
            "UI",
        ),
    ],
)

_add(
    "BRS-SSI-005",
    [
        (
            "[Positive][Spleen] Perform ordered workflow: 2D Brush → ROI SpleenFS → assign first slice → ROI SpleenLS → assign last slice; verify eCRF shows slice numbers and LVD calculated.",
            "Spleen first/last slice numbers displayed; LVD calculated correctly.",
            "Positive",
        ),
        (
            "[Negative][Spleen] Incomplete workflow (only first slice ROI created) and check LVD/capture.",
            "LVD not fully calculated / incomplete capture until both first and last slice assigned.",
            "Negative",
        ),
        (
            "[UI][Spleen] Verify SpleenFS/SpleenLS naming and eCRF display of slice numbers and LVD fields.",
            "Controls SpleenFirstSlice/SpleenLastSlice and LVD display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-SSI-035",
    [
        (
            "[Positive][Spleen] After valid spleen slices, verify hidden fields store First/Last/NextToLast locations and SliceThickness.",
            "Hidden fields populated: SpleenFirstSliceLocation, SpleenLastSliceLocation, SpleenNextToLastSliceLocation, SpleenSliceThickness.",
            "Positive",
        ),
        (
            "[Negative][Spleen] When LVD = n/a, verify hidden location/thickness fields are blank.",
            "Hidden fields blank when LVD = n/a.",
            "Negative",
        ),
        (
            "[UI][Spleen] Confirm location/thickness fields are hidden from reader UI but stored in database.",
            "Fields not visible on eCRF; values present in DB/export.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S1QA-000a",
    [
        (
            "[Positive][Session 1] TAA Optimal/Readable → answer IndicateFDGAvidity FDG-Avid; add required PET Positive P1; sign off.",
            "FDG-avidity saved; sign-off succeeds with P1 present.",
            "Positive",
        ),
        (
            "[Positive][Session 1] TAA Not Readable or no PET at screening → IndicateFDGAvidity defaults to NE.",
            "FDG-avidity auto-populated NE.",
            "Positive",
        ),
        (
            "[Negative][Session 1] Attempt to add PET/CT ROI before answering FDG-avidity.",
            "Error: FGD-Avidity answer is required. ROI not added.",
            "Negative",
        ),
        (
            "[Negative][Session 1] FDG-Avid selected but no PET Positive lesion; attempt sign-off.",
            "Error: FDG-Avid subjects require PET Positive lesion to be added. Sign-off blocked.",
            "Negative",
        ),
        (
            "[UI][Session 1] Verify IndicateFDGAvidity options FDG-Avid / Non FDG-Avid / NE (NE non-selectable) and field enablement rules.",
            "Options and enable/disable/default states match requirement.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S1QA-002",
    [
        (
            "[Positive][Session 1] TAA Optimal → select bone marrow No evidence / Diffuse benign / Focal positive as applicable.",
            "EvidenceBoneMarrow saved; one answer required.",
            "Positive",
        ),
        (
            "[Positive][Session 1] TAA Not Readable or no PET → auto-populate and disable Not Evaluable (PET not available or very bad quality).",
            "Field auto NE and disabled.",
            "Positive",
        ),
        (
            "[Negative][Session 1] Non FDG-avid + Focal FDG avid disease in bone marrow (positive) at sign-off.",
            "Message: Subjects deemed Non FDG-avid may not have a positive bone marrow assessment... Change required; sign-off blocked until corrected.",
            "Negative",
        ),
        (
            "[UI][Session 1] Verify bone marrow answer options and disabled auto-NE state.",
            "UI options and disable behavior match specification.",
            "UI",
        ),
    ],
)

_add(
    "BRS-LM-035",
    [
        (
            "[Positive][Session 2] Index lesion decreases to 4×4 mm → TSTM applied as 5×5 mm for area calculation.",
            "TSTM status; Final LDi/SDi recorded as 5×5 mm.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Lesion decreases to 10×4 mm → only SDi axis recorded as 5 mm (10×5).",
            "Only below-threshold axis set to 5 mm.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Split fragment below 5 mm receives TSTM 5 mm assignment.",
            "Fragment below threshold assigned TSTM 5 mm.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Confirm true <5 mm values are not used for PPD when TSTM applies (except LM-036 Yes).",
            "PPD uses defaulted 5 mm axes, not raw <5 values.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify TSTM / Present and Normal (TSTM) status display on lesion panel.",
            "Status and defaulted measurements visible correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-LM-036",
    [
        (
            "[Positive][Session 2] Diameter >0 and <5 mm; ReliablyMeasured = Yes → true LDi/SDi applied; Present and Normal (nodal) / Present (extranodal).",
            "True measurements retained; status Present and Normal / Present without TSTM defaults.",
            "Positive",
        ),
        (
            "[Positive][Session 2] ReliablyMeasured = No → TSTM defaults of 5 mm applied per BRS-LM-035 matrix.",
            "Default LDi/SDi = 5 where applicable; ROI Status Present and Normal (TSTM) / Present (TSTM).",
            "Positive",
        ),
        (
            "[Negative][Session 2] Leave ReliablyMeasured unanswered when diameters in (0,5) and attempt continue.",
            "Question required; user cannot proceed without Yes/No.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify ReliablyMeasured Yes/No context menu appears only when threshold >0 and <5 mm.",
            "Control appears only for applicable diameter range.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CICR-001",
    [
        (
            "[Positive][CT Index CR] All extranodal Index resolved/Not Present and all nodal Index LDi ≤15 mm.",
            "CT Index Response = CR.",
            "Positive",
        ),
        (
            "[Negative][CT Index CR] One nodal Index LDi = 16 mm with extranodal resolved.",
            "CT Index Response is not CR.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify CR displayed on Target Response / TPS panel.",
            "CR shown correctly in UI.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CIPR-001",
    [
        (
            "[Positive][CT Index PR] SPD decrease ≥50% from screening and not CR/PD/NE.",
            "CT Index Response = PR.",
            "Positive",
        ),
        (
            "[Negative][CT Index PR] SPD decrease 49% from screening (CR/PD/NE not met).",
            "CT Index Response is not PR (falls to SD or other).",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify % SPD change and PR response on tracking/TPS panels.",
            "Values and PR label display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CISD-001",
    [
        (
            "[Positive][CT Index SD] Criteria for CR, PR, PD, NE, NA not met.",
            "CT Index Response = SD.",
            "Positive",
        ),
        (
            "[Negative][CT Index SD] Case meeting PR (≥50% SPD decrease) is not labeled SD.",
            "Response = PR, not SD.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify SD appears on response fields when applicable.",
            "SD displayed correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CIPD-001",
    [
        (
            "[Positive][CT Index PD] Single lesion meets PD while sum of lesions would not → PD derived from single lesion.",
            "CT Index Response = PD based on single lesion instance.",
            "Positive",
        ),
        (
            "[Negative][CT Index PD] No individual lesion meets PD conditions though sum changes modestly.",
            "PD not derived from sums alone.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify PD and contributing lesion highlighted/available for review.",
            "PD shown; lesion-level data visible.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CIPD-005",
    [
        (
            "[Positive][CT Index PD] Lesion LDi >15, ≥50% PPD↑ from nadir, and absolute ↑ meets 5 mm (nadir ≤20) or 10 mm (nadir >20).",
            "CT Index Response = PD.",
            "Positive",
        ),
        (
            "[Positive][CT Index PD] Nadir PPD = 0 and current PPD numeric non-zero → condition 2 true.",
            "Condition 2 considered met; PD if other conditions satisfied.",
            "Positive",
        ),
        (
            "[Positive][CT Index PD] Merged lesions: combined PPD used for nadir; when abnormal + 50% met, LDi/SDi increase (#3) not required; #3 no longer applies subsequently.",
            "PD derived per merge rules; subsequent #3 waived.",
            "Positive",
        ),
        (
            "[Negative][CT Index PD] Lesion meets 50% PPD increase but LDi ≤15 mm (not abnormal).",
            "PD not assigned.",
            "Negative",
        ),
        (
            "[Negative][CT Index PD] Abnormal + 50% PPD but absolute LDi/SDi increase below threshold (non-merged).",
            "PD not assigned.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify nadir PPD, % change, and absolute change fields used for PD review.",
            "Calculation fields display with correct rounding and support PD verification.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CIPD-010",
    [
        (
            "[Positive][CT Index PD] Extranodal lesion resolved then later Present with LDi >10 mm.",
            "CT Index Response = PD.",
            "Positive",
        ),
        (
            "[Negative][CT Index PD] Reappeared extranodal lesion LDi = 10 mm (not >10).",
            "PD not triggered by this rule.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify status history Resolved → Present and PD response display.",
            "Status transition and PD visible.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CINE-001",
    [
        (
            "[Positive][CT Index NE] Current TP Not Readable and Index lesions present at screening.",
            "CT Index Response = NE.",
            "Positive",
        ),
        (
            "[Negative][CT Index NE] TP Readable with evaluable Index lesions → not forced NE by this rule.",
            "Response not NE solely from this rule.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify NE displayed when TP Not Readable.",
            "NE shown on Target Response.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CINE-005",
    [
        (
            "[Positive][CT Index NE] At least one Index PPD not evaluable and PD criteria not met.",
            "CT Index Response = NE.",
            "Positive",
        ),
        (
            "[Negative][CT Index NE] One lesion NE but another lesion meets full PD criteria.",
            "PD takes precedence; response = PD not NE.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify NE lesion status and Index Response NE.",
            "UI reflects NE lesion and response.",
            "UI",
        ),
    ],
)

_add(
    "BRS-CINA-001",
    [
        (
            "[Positive][CT Index NA] No Index/Target at screening and current TP not Not Readable.",
            "CT Index Response = NA.",
            "Positive",
        ),
        (
            "[Negative][CT Index NA] No Index at screening but current TP Not Readable → not NA via this rule (covered by NE rules).",
            "NA not incorrectly assigned when Not Readable.",
            "Negative",
        ),
        (
            "[UI][CT Index] Verify NA displayed on Target Response.",
            "NA visible in UI.",
            "UI",
        ),
    ],
)

# Spleen specialized compact sets
for rid, pos, neg, ui_e in [
    (
        "BRS-SR-000",
        "[Positive][Spleen] Verify Spleen Nadir = smallest prior LVD excluding current TP.",
        "[Negative][Spleen] Confirm current TP LVD is not included in nadir.",
        "[UI][Spleen] Nadir LVD value visible/used in spleen response panel.",
    ),
    (
        "BRS-SR-001",
        "[Positive][Spleen] Screening & TP1 unevaluable; TP2 first measurable → TP2 becomes Reference for subsequent responses.",
        "[Negative][Spleen] Later TP does not incorrectly use unevaluable screening as Reference when TP2 is first assessable.",
        "[UI][Spleen] Reference TP indication/display for spleen response.",
    ),
    (
        "BRS-SR-002",
        "[Positive][Spleen] LVD 131 mm = abnormal; LVD 130 mm = normal.",
        "[Negative][Spleen] LVD 130 mm not treated as abnormal.",
        "[UI][Spleen] Abnormal/Normal designation shown from LVD threshold.",
    ),
    (
        "BRS-SR-003",
        "[Positive][Spleen] Edit Enlarged SpleenAssessment to Normal; designation Normal overrides measurement.",
        "[Negative][Spleen] After Normal override, measurement >130 does not force Abnormal designation.",
        "[UI][Spleen] SpleenAssessment editable Enlarged→Normal control.",
    ),
    (
        "BRS-SRCR-001",
        "[Positive][Spleen CR] Abnormal at Reference (ASV>0) regressed to LVD ≤130 → CR.",
        "[Negative][Spleen CR] Remains LVD >130 → not CR.",
        "[UI][Spleen] Spleen Response displays CR.",
    ),
    (
        "BRS-SRCR-002",
        "[Positive][Spleen CR] Change Spleen Status Enlarged→Normal when prior abnormal at nadir → CR.",
        "[Negative][Spleen CR] Leave Enlarged with still abnormal size → not CR via this rule.",
        "[UI][Spleen] Status change reflects CR on panel.",
    ),
    (
        "BRS-SRPR-001",
        "[Positive][Spleen PR] Abnormal remains abnormal; ASV ↓ >50% from Reference → PR.",
        "[Negative][Spleen PR] ASV ↓ 50% or less without PD → not PR.",
        "[UI][Spleen] PR displayed for spleen response.",
    ),
    (
        "BRS-SRSD-001",
        "[Positive][Spleen SD] Abnormal remains abnormal; neither PR nor PD → SD.",
        "[Negative][Spleen SD] Meets PR → not SD.",
        "[UI][Spleen] SD displayed correctly.",
    ),
    (
        "BRS-SRPD-020",
        "[Positive][Spleen PD] Normal at nadir → abnormal LVD>130 and ↑ ≥20 mm vs nadir → PD.",
        "[Negative][Spleen PD] Becomes abnormal but LVD increase 19 mm → not PD by this rule.",
        "[UI][Spleen] PD displayed for spleen.",
    ),
    (
        "BRS-SRPD-025",
        "[Positive][Spleen PD] Abnormal at nadir; ASV ↑ >50% and LVD ↑ ≥10 mm → PD.",
        "[Negative][Spleen PD] ASV ↑ >50% but LVD ↑ 9 mm → not PD.",
        "[UI][Spleen] PD and related LVD/ASV values visible.",
    ),
    (
        "BRS-SRNE-001",
        "[Positive][Spleen NE] TP Not Readable → Spleen Response NE.",
        "[Negative][Spleen NE] Readable TP with measurable spleen → not NE by this rule.",
        "[UI][Spleen] NE shown when TP Not Readable.",
    ),
    (
        "BRS-SRNE-005",
        "[Positive][Spleen NE] Status Not Present due to technical issues / images not available → NE.",
        "[Negative][Spleen NE] Present measurable spleen → not NE.",
        "[UI][Spleen] Status options and NE response display.",
    ),
    (
        "BRS-SRNE-010",
        "[Positive][Spleen NE] Prior measurement exists; current Not Present, Surgery → NE.",
        "[Negative][Spleen NE] Not Present, Surgery at screening only path covered by NC rules — verify not double-classified incorrectly.",
        "[UI][Spleen] Not Present, Surgery status and NE response.",
    ),
    (
        "BRS-SRNE-015",
        "[Positive][Spleen NE] Screening unevaluable; first on-study LVD >130 → NE.",
        "[Negative][Spleen NE] Screening unevaluable; first on-study ≤130 → NC not NE.",
        "[UI][Spleen] NE displayed for first abnormal on-study after unevaluable screening.",
    ),
    (
        "BRS-SRNC-001",
        "[Positive][Spleen NC] Normal at Reference remains normal → NC.",
        "[Negative][Spleen NC] Remains normal but meets other response incorrectly → stays NC.",
        "[UI][Spleen] NC displayed.",
    ),
    (
        "BRS-SRNC-005",
        "[Positive][Spleen NC] Screening unevaluable; first on-study normal ≤130 (or Enlarged→Normal) → NC.",
        "[Negative][Spleen NC] First on-study abnormal >130 → not NC (NE per SRNE-015).",
        "[UI][Spleen] NC displayed.",
    ),
    (
        "BRS-SRNC-010",
        "[Positive][Spleen NC] Normal at Reference becomes abnormal but not PD → NC.",
        "[Negative][Spleen NC] Meets PD thresholds → PD not NC.",
        "[UI][Spleen] NC vs PD distinction visible.",
    ),
    (
        "BRS-SRNC-015",
        "[Positive][Spleen NC] Not Present, surgery at screening → NC.",
        "[Negative][Spleen NC] Not Present, surgery only on-study after prior measurement → NE not NC.",
        "[UI][Spleen] NC for screening surgery absence.",
    ),
    (
        "BRS-SRNC-020",
        "[Positive][Spleen NC] Screening unevaluable and first assessable on-study Not Present, surgery → NC.",
        "[Negative][Spleen NC] Distinguish from SRNE-010 prior measurement + surgery path.",
        "[UI][Spleen] NC displayed for this path.",
    ),
]:
    _add(
        rid,
        [
            (pos, "Spleen response/logic matches requirement.", "Positive"),
            (neg, "Incorrect alternate classification is not applied.", "Negative"),
            (ui_e, "UI displays spleen status/response per requirement.", "UI"),
        ],
    )

_add(
    "BRS-S2DD-025",
    [
        (
            "[Positive][Session 2] Display Most FDG-avid P1 Lesion, SUVmax, %Change Screening, %Change Nadir with correct formulas.",
            "P1 panel values calculate and display correctly.",
            "Positive",
        ),
        (
            "[Positive][Session 2] PET lesion identified but no SUV values → auto-populate NE.",
            "SUVmax / % change fields show NE.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Verify %ΔSUVmax uses worst-case lesion per TP and nadir comparison for PMD config.",
            "Wrong lesion or wrong baseline not used in % change.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify P1 informational panel fields and control names display.",
            "WorstLesionsSUVmax, PerChangeSUVmaxFromScreening, PerChangeSUVmaxFromNadir visible as specified.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-002",
    [
        (
            "[Positive][Session 2] Select bone marrow options including New or recurrent FDG avid foci when applicable.",
            "Answer saved; one answer required.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Select Focal positive when prior TP was negative marrow → warning to consider new/recurrent; change required.",
            "Message displayed; reader must modify selection.",
            "Negative",
        ),
        (
            "[Negative][Session 2] Select New or recurrent positive while P1 PET Assessment Negative → conflict message; change required.",
            "Message displayed; modification required.",
            "Negative",
        ),
        (
            "[Negative][Session 2] Sign-off with Focal positive marrow and PET TPR CMR → conflict message; change required.",
            "Sign-off blocked until bone marrow or P1 assessment reconsidered.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify additional New/recurrent option and auto-NE disable states.",
            "Options and disabled NE state match Session 2 specification.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-010b",
    [
        (
            "[Positive][Session 2] FDG-positive at screening + 5PS 1/2/3 → DeterminationChange auto Decrease from screening and disabled.",
            "Auto-populated Decrease and disabled.",
            "Positive",
        ),
        (
            "[Positive][Session 2] New CT lesion PET Assessment Positive → auto Increase from nadir disabled.",
            "Auto Increase applied and disabled.",
            "Positive",
        ),
        (
            "[Positive][Session 2] PET NE at screening + 5PS 4/5 with prior assessable PET → options limited to Increase from nadir and NE.",
            "Only allowed options enabled.",
            "Positive",
        ),
        (
            "[Negative][Session 2] 5PS NE → DeterminationChange auto NE disabled; reader cannot select Increase/Decrease.",
            "Field locked to NE.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify DeterminationChange options and auto-populate/disable visual states.",
            "UI states match auto-populate rules.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-020",
    [
        (
            "[Positive][Session 2] PET TPR auto-populates CMR/PMR/NMR/PMD/NA/NE per staging table and bone marrow new/recurrent → PMD.",
            "TimePointResponsePET correct and read-only.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Prior PMD → subsequent PET TPR locked PMD.",
            "PMD populated and disabled.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Prior CMR and current derived not CMR/PMD → NE; prior CMR/PMR and current NMR → NE.",
            "Out-of-sequence responses yield NE.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify PET TPR field always read-only with correct option list.",
            "Field disabled; values display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-022",
    [
        (
            "[Positive][Session 2] Designated PET defaults to PET TPR; editable for FDG-avid when not disabled; prior CMR limits options to CMR/PMD/NE.",
            "Defaults and option limits correct.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Non FDG-avid: options limited to PMD/NA/NE.",
            "Only Non FDG-avid options available.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Designated PET disabled for NE (Not Readable / no PET / screening NR / out-of-sequence) and for PMD.",
            "Field not editable in disabled cases.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify enable/disable and dropdown contents for FDG vs Non FDG-avid.",
            "UI option sets and states match requirement.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-021",
    [
        (
            "[Positive][Session 2] Designated PET ≠ PMD → PET TPR Date = latest imaging date; = PMD → earliest date.",
            "Hidden date populated per rule.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Not Readable due to Time Point missing → date n/a.",
            "Date = n/a.",
            "Negative",
        ),
        (
            "[UI][Session 2] Confirm TimePointResponsePETDate is hidden from reader.",
            "Field not visible on eCRF; value stored.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-023",
    [
        (
            "[Positive][Session 2] Change Designated PET so it differs from PET TPR → comment enabled and completed.",
            "CommentChangingPETResponse required and saved.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Designated ≠ PET TPR with blank comment → sign-off blocked.",
            "System requires comment when enabled.",
            "Negative",
        ),
        (
            "[UI][Session 2] Comment field hidden/disabled when Designated matches PET TPR; enabled when mismatch.",
            "Enablement matches match/mismatch state.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-045",
    [
        (
            "[Positive][Session 2] CT TPR auto-populates from Target/Non-Target/New lesions/Spleen per TPR table.",
            "TimePointResponseCT correct and read-only.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Reader-choice rows 6/11/14 enable PR|NE or SD|NE selection as specified.",
            "Reader can choose allowed options; selection saved.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Prior CR with current not CR/PD → NE; prior CR/PR with current SD → NE (overrides PR/SD).",
            "Out-of-sequence NE applied.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify CT TPR read-only display and conditional choice dropdowns.",
            "UI states match table rows and NE override rules.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-050",
    [
        (
            "[Positive][Session 2] PET+CT TPR auto-populates from CT TPR + PET TPR combination rules.",
            "TimePointResponseCTPET correct.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Technical adequacy Not Readable or PET out-of-sequence NE → PET+CT TPR NE.",
            "PET+CT TPR = NE.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify PET+CT TPR read-only field and option list including NED.",
            "Field disabled; options display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-052",
    [
        (
            "[Positive][Session 2] Designated PET+CT defaults to PET+CT TPR; prior CR limits to CR/PD/NE; prior PR to CR/PR/PD/NE.",
            "Defaults and limits correct for FDG-avid.",
            "Positive",
        ),
        (
            "[Positive][Session 2] Non FDG-avid dropdowns based on derived NED/CR/PR/SD as specified.",
            "Non FDG-avid option sets correct.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Disabled when NE (Not Readable / screening NR / out-of-sequence) or PD.",
            "Field not editable.",
            "Negative",
        ),
        (
            "[UI][Session 2] Verify Designated PET+CT enablement and option lists.",
            "UI matches avidity and prior-response rules.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-051",
    [
        (
            "[Positive][Session 2] Designated PET+CT ≠ PD → latest date; = PD → earliest date.",
            "Hidden PET+CT date correct.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Time Point missing Not Readable → date n/a.",
            "Date = n/a.",
            "Negative",
        ),
        (
            "[UI][Session 2] Confirm TimePointResponseCTPETDate hidden.",
            "Field hidden; stored correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S2QA-053",
    [
        (
            "[Positive][Session 2] Designated PET+CT ≠ PET+CT TPR → comment required and saved.",
            "CommentChangingCTPETResponse required when enabled.",
            "Positive",
        ),
        (
            "[Negative][Session 2] Mismatch with empty comment → blocked.",
            "Sign-off/continue blocked until comment entered.",
            "Negative",
        ),
        (
            "[UI][Session 2] Comment enable only on mismatch.",
            "UI enablement correct.",
            "UI",
        ),
    ],
)

_add(
    "BRS-REA-001",
    [
        (
            "[Positive][Reassessment] Click reassessment for a follow-up TP; Session 1/2 answers shown; FDG-avidity and ImageAdequacy locked; Baseline READ ONLY.",
            "Reassessment window opens with locked baseline avidity/TAA; baseline remains read-only.",
            "Positive",
        ),
        (
            "[Negative][Reassessment] Attempt to change IndicateFDGAvidity or ImageAdequacy during reassessment.",
            "Fields locked; changes not allowed.",
            "Negative",
        ),
        (
            "[UI][Reassessment] Verify reassessment button and locked field visual state.",
            "Reassessment UI and lock indicators correct.",
            "UI",
        ),
    ],
)

_add(
    "BRS-QARA-010",
    [
        (
            "[Positive][Reassessment] TP with PD/PMD → Pseudoprogression question visible; Yes → PD/PMD responses change to NE and recalculate; only PD/PMD change.",
            "Pseudoprogression Yes converts PD/PMD to NE and recalculates related TPRs per rules.",
            "Positive",
        ),
        (
            "[Positive][Reassessment] Pseudoprogression No with once-PD-always-PD → recalculate CT/PET per S2QA-045/020.",
            "Underlying responses recalculated.",
            "Positive",
        ),
        (
            "[Negative][Reassessment] Question not visible when TPR is not PD/PMD.",
            "ReassessPseudoprgression hidden.",
            "Negative",
        ),
        (
            "[UI][Reassessment] Verify question default No, Yes/No options, and visibility only for PD/PMD.",
            "UI visibility and defaults match requirement.",
            "UI",
        ),
    ],
)

# Session 3 display / QA specialized brief sets for key interactive ones
_add(
    "BRS-S3QA-010",
    [
        (
            "[Positive][Session 3] Answer OnstudyRadiotherapySurgicalPresent Yes/No; re-assess/insert maintains prior Global answer.",
            "Answer required and maintained across re-assess/insert.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Attempt sign-off without answering Yes/No.",
            "Blocked; one answer required.",
            "Negative",
        ),
        (
            "[UI][Session 3] Verify question always displayed with Yes/No.",
            "Field always visible; options correct.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S3QA-011",
    [
        (
            "[Positive][Session 3] S3QA-010=Yes enables OnstudyDataImpactCT; default No; editable.",
            "Field enabled with default No.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Answer Yes without CT lesion overrides NE radiotherapy at sign-off → error requiring overrides or change to No.",
            "Sign-off blocked with specified message.",
            "Negative",
        ),
        (
            "[Negative][Session 3] Answer No while CT lesion overrides present → error requiring removal or change to Yes.",
            "Sign-off blocked with specified message.",
            "Negative",
        ),
        (
            "[UI][Session 3] Field disabled when S3QA-010=No; enabled when Yes.",
            "Enablement tied to S3QA-010.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S3QA-012",
    [
        (
            "[Positive][Session 3] S3QA-010=Yes enables PET impact question with Yes/No/Yes: with no change required.",
            "Options available; default No.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Answer Yes without required PET TPR override → sign-off message; Answer No with PET override present → sign-off message.",
            "Sign-off blocked with specified messages.",
            "Negative",
        ),
        (
            "[UI][Session 3] Verify three answer options and enablement.",
            "UI matches specification.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S3QA-014",
    [
        (
            "[Positive][Session 3] With RT present & impacts CT, override Target lesion status Yes → reason Not evaluable – on study radiotherapy; carries to subsequent TPs.",
            "Override applied and carried forward; later TP context menu suppressed for that lesion.",
            "Positive",
        ),
        (
            "[Positive][Session 3] Subsequent Global: undo override Yes→No at earliest TP → message and required comment.",
            "Message shown; Reason for Override Lesion Status Change required.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Override not allowed when S3QA-010/011 not both Yes.",
            "Override context menu not available.",
            "Negative",
        ),
        (
            "[UI][Session 3] Verify Target panel override dialog Yes/No and comment box visibility rules.",
            "UI dialogs and comment visibility correct.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S3QA-017",
    [
        (
            "[Positive][Session 3] Override Non-Target Present/Present and Normal lesion; do not change Present with progression; carry forward.",
            "Override applied per rules; Present with progression unchanged.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Attempt override when lesion not Present/Present and Normal or when impact flags not Yes.",
            "Override not offered.",
            "Negative",
        ),
        (
            "[UI][Session 3] Non-Target override dialog and carry-forward display.",
            "UI behaves per requirement.",
            "UI",
        ),
    ],
)

_add(
    "BRS-S3QA-024",
    [
        (
            "[Positive][Session 3] Override Designated PET TPR Yes with reason (RT/surgery/both) when impact PET=Yes and Designated ≠ PMD; carry to subsequent non-PMD TPs.",
            "Override applied with reason; PMD TPs unchanged; later menus suppressed.",
            "Positive",
        ),
        (
            "[Positive][Session 3] Undo override in later Global at earliest TP → comment required.",
            "Comment box required after message.",
            "Positive",
        ),
        (
            "[Negative][Session 3] Cannot override when Designated PET is PMD or impact flags not both Yes.",
            "Override not available.",
            "Negative",
        ),
        (
            "[UI][Session 3] Override PET TPR dialogs and reason options display.",
            "UI matches requirement.",
            "UI",
        ),
    ],
)

_add(
    "BRS-OS5QA-070",
    [
        (
            "[Positive][Session 5] Clinical assessment after last radiology = Yes → clinical-only TP created with required Overall PET/CT/PET+CT/Designated responses, Clinical Data, Exam Date.",
            "Clinical-only time point created and all required fields enforced.",
            "Positive",
        ),
        (
            "[Negative][Session 5] Leave required clinical-only TP fields blank and attempt complete.",
            "System blocks until required fields completed.",
            "Negative",
        ),
        (
            "[UI][Session 5] Verify clinical-only TP naming (next consecutive / Timepoint 2 if screening-only) and n/a min/max dates.",
            "Naming and n/a dates display correctly.",
            "UI",
        ),
    ],
)

_add(
    "BRS-OS5DD-061",
    [
        (
            "[Positive][Session 5] Open Reference Notes panel and verify Appendix A content; toggle off/on.",
            "Panel displays Appendix A; toggle works.",
            "Positive",
        ),
        (
            "[UI][Session 5] Verify panel name Reference Notes and toggle control.",
            "Panel labeled correctly; toggle visible.",
            "UI",
        ),
        (
            "[Negative][Session 5] Toggle off → content hidden; assessments still completable.",
            "Hidden panel does not block workflow.",
            "Negative",
        ),
    ],
)

_add(
    "BRS-OS5DD-062",
    [
        (
            "[Positive][Session 5] Open PET + CT + Clinical Assessment panel (Appendix B); toggle off/on.",
            "Panel displays Appendix B; toggle works.",
            "Positive",
        ),
        (
            "[UI][Session 5] Verify panel name and toggle.",
            "Correct label and toggle behavior.",
            "UI",
        ),
        (
            "[Negative][Session 5] Toggle off does not remove required assessment fields.",
            "Informational panel hide does not remove eCRF required fields.",
            "Negative",
        ),
    ],
)


def build_rows() -> tuple[list[dict], list[dict]]:
    test_rows: list[dict] = []
    trace_rows: list[dict] = []
    ts_num = 1

    for req_id, desc, module in REQUIREMENTS:
        sc_list = scenarios_for(req_id, desc, module)
        for scenario, expected, _case_type in sc_list:
            ts_id = f"TS-{ts_num:04d}"
            ts_num += 1
            test_rows.append(
                {
                    "Requirement ID": req_id,
                    "Test Scenario ID": ts_id,
                    "Test Scenario": scenario,
                    "Expected result": expected,
                }
            )
            trace_rows.append(
                {
                    "Requirement ID": req_id,
                    "Requirement Description": desc,
                    "Test Scenario ID": ts_id,
                    "Status": "Not Executed",
                }
            )
    return test_rows, trace_rows


def style_header(ws, fill_color: str) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor=fill_color)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin


def autosize(ws, widths: dict[int, int]) -> None:
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = Border(
                left=Side(style="thin", color="D0D0D0"),
                right=Side(style="thin", color="D0D0D0"),
                top=Side(style="thin", color="D0D0D0"),
                bottom=Side(style="thin", color="D0D0D0"),
            )


def main() -> None:
    test_rows, trace_rows = build_rows()
    wb = Workbook()

    # Sheet 1 — Test Scenarios
    ws1 = wb.active
    ws1.title = "Test Scenarios"
    headers1 = ["Requirement ID", "Test Scenario ID", "Test Scenario", "Expected result"]
    ws1.append(headers1)
    for row in test_rows:
        ws1.append([row[h] for h in headers1])
    style_header(ws1, "0B6E6E")
    autosize(ws1, {1: 18, 2: 16, 3: 85, 4: 55})
    ws1.auto_filter.ref = ws1.dimensions
    ws1.freeze_panes = "A2"
    ws1.row_dimensions[1].height = 22

    # Sheet 2 — Traceability Matrix
    ws2 = wb.create_sheet("Traceability Matrix")
    headers2 = [
        "Requirement ID",
        "Requirement Description",
        "Test Scenario ID",
        "Status",
    ]
    ws2.append(headers2)
    for row in trace_rows:
        ws2.append([row[h] for h in headers2])
    style_header(ws2, "0A4F6E")
    autosize(ws2, {1: 18, 2: 90, 3: 16, 4: 16})
    ws2.auto_filter.ref = ws2.dimensions
    ws2.freeze_panes = "A2"
    ws2.row_dimensions[1].height = 22

    # Sheet 3 — Requirements catalog (for auditor convenience)
    ws3 = wb.create_sheet("Requirements Catalog")
    ws3.append(["Requirement ID", "Requirement Description", "Module / Session"])
    for req_id, desc, module in REQUIREMENTS:
        ws3.append([req_id, desc, module])
    style_header(ws3, "334155")
    autosize(ws3, {1: 18, 2: 100, 3: 32})
    ws3.auto_filter.ref = ws3.dimensions
    ws3.freeze_panes = "A2"

    # Cover / summary
    ws0 = wb.create_sheet("Cover", 0)
    ws0["A1"] = "Lugano Criteria Windows Application — Test Scenarios & Traceability Matrix"
    ws0["A1"].font = Font(bold=True, size=14, color="0B6E6E")
    ws0["A3"] = "Source document"
    ws0["B3"] = "General Requirements (BR-S-DR) — R1979-ONC-2105 / BioREAD Lugano"
    ws0["A4"] = "Sessions covered"
    ws0["B4"] = "Session 1 (Baseline), Session 2 (Follow-up), Session 3 (Global), Session 4 (Adjudication), Session 5 (Clinical), plus cross-session modules"
    ws0["A5"] = "Case types"
    ws0["B5"] = "Positive, Negative, UI"
    ws0["A6"] = "Test Scenario ID format"
    ws0["B6"] = "TS-0001"
    ws0["A7"] = "Total requirements"
    ws0["B7"] = len(REQUIREMENTS)
    ws0["A8"] = "Total test scenarios"
    ws0["B8"] = len(test_rows)
    ws0["A9"] = "Traceability status default"
    ws0["B9"] = "Not Executed"
    ws0["A11"] = "Sheets"
    ws0["B11"] = "Cover | Test Scenarios | Traceability Matrix | Requirements Catalog"
    ws0.column_dimensions["A"].width = 28
    ws0.column_dimensions["B"].width = 110

    out = Path("/workspace/validation/test_artifacts/Lugano_BRS_Test_Scenarios_Traceability.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}")
    print(f"Requirements: {len(REQUIREMENTS)}")
    print(f"Test scenarios: {len(test_rows)}")
    print(f"Traceability rows: {len(trace_rows)}")


if __name__ == "__main__":
    main()
