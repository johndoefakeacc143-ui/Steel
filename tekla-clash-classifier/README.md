# Tekla Clash Classifier

Classifies Tekla `ClashCheck.log` GUID pairs using a Universal Report (parts) and a Bolt checking report, then writes:

- `Part_to_Part_Clash`
- `Part_to_Bolt_Clash`
- `Bolt_to_Bolt_Clash`
- `Duplicate members` (log description contains `inside member`)

## Quick start (Python)

```bash
python3 scripts/classify_clashes.py \
  --universal /path/to/Universal_Report.xlsx \
  --bolt /path/to/BOLT_CHECKING_REPORT.xlsx \
  --clash /path/to/ClashCheck.log \
  --output classified_clashes.xlsx
```

## Excel macro workbook

```bash
python3 scripts/build_macro_workbook.py --out-dir dist
```

1. Open `dist/tekla-clash-classifier.xlsx`, import `TeklaClashClassifier.bas`, Save As `.xlsm`.
2. Run `SetupClassifierUI` once (or reopen) to place the **RUN** Form Control button.
3. Click **RUN** → pick Universal / Bolt / ClashCheck.log → SaveAs a **new** results workbook.

See Project store howto: `docs/tekla-clash-macro-howto.md`.

## Validate against sample inputs

```bash
python3 scripts/validate_against_sample.py
```
