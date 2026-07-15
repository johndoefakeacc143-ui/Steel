# Lugano BRS Test Scenarios & Traceability

Validation artifacts for the Lugano criteria BioREAD windows application (protocol **R1979-ONC-2105**), derived from *General Requirements (BR-S-DR)*.

## Deliverable

**[`validation/test_artifacts/Lugano_BRS_Test_Scenarios_Traceability.xlsx`](validation/test_artifacts/Lugano_BRS_Test_Scenarios_Traceability.xlsx)**

| Sheet | Columns |
|---|---|
| **Test Scenarios** | Requirement ID \| Test Scenario ID \| Test Scenario \| Expected result |
| **Traceability Matrix** | Requirement ID \| Requirement Description \| Test Scenario ID \| Status |
| **Requirements Catalog** | All copied requirement IDs/descriptions by module/session |
| **Cover** | Summary counts and scope |

- Test Scenario IDs: `TS-0001` …  
- Case types: **Positive**, **Negative**, **UI**  
- Sessions covered: 1 (Baseline), 2 (Follow-up), 3 (Global), 4 (Adjudication), 5 (Clinical), plus cross-session modules (TAA, lesions, spleen, tracking panels, reassessment)

## Regenerate

```bash
python scripts/generate_lugano_test_scenarios_excel.py
```

## Source

`validation/test_artifacts/General_Requirements_BR-S-DR.txt`
