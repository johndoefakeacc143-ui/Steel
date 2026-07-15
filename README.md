# BioREAD BRS Test Scenarios & Traceability

Validation Excel workbooks for BioREAD windows applications, derived from provided General Requirements documents.

## Deliverables

| Study / Source | Excel workbook |
|---|---|
| Lugano (BR-S-DR / R1979-ONC-2105) | [`validation/test_artifacts/Lugano_BRS_Test_Scenarios_Traceability.xlsx`](validation/test_artifacts/Lugano_BRS_Test_Scenarios_Traceability.xlsx) |
| CLL/SLL (REq_dd36) | [`validation/test_artifacts/CLLSLL_BRS_Test_Scenarios_Traceability.xlsx`](validation/test_artifacts/CLLSLL_BRS_Test_Scenarios_Traceability.xlsx) |

### Sheet layout (both workbooks)

| Sheet | Columns |
|---|---|
| **Test Scenarios** | Requirement ID \| Test Scenario ID \| Test Scenario \| Expected result |
| **Traceability Matrix** | Requirement ID \| Requirement Description \| Test Scenario ID \| Status |
| **Requirements Catalog** | Requirement IDs / descriptions by module/session |
| **Cover** | Summary counts and scope |

- Test Scenario IDs: `TS-0001` …
- Case types: **Positive**, **Negative**, **UI**
- Traceability Status default: `Not Executed`

## Regenerate

```bash
python scripts/generate_lugano_test_scenarios_excel.py
python scripts/generate_cllsll_test_scenarios_excel.py
```

## Sources

- `validation/test_artifacts/General_Requirements_BR-S-DR.txt`
- `validation/test_artifacts/General_Requirements_CLLSLL_REq.txt`
