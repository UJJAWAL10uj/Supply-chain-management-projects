# Localization Gap and Supplier Risk Map: Mobile Aluminium Enclosures in India

## Question
Which parts of the mobile aluminium-enclosure value chain can India make today versus import, which Indian
suppliers carry the most risk, and what should a sourcing team do about the top three risks?

## Deliverables
| File | What it is |
|---|---|
| `Enclosure_Localization_Supplier_Risk_Dashboard.xlsx` | Dashboard, value-chain gap map (9 stages), 10 scored suppliers, risk register with heat map, mitigation plan, 26 cited sources. About 200 formulas, 0 errors |
| `Risk_Mitigation_Plan_Enclosure_Supply_Chain.pdf` | Two-page risk-mitigation plan (numbers and actions read from the workbook) |
| `build_p2.py`, `build_p2_plan.py` | Regenerate the workbook and the plan |

## Method
- Supplier risk = weighted average of four 1-5 scores (single-source exposure, capacity and ramp, geopolitical / China-linked exposure, quality readiness); weights editable.
- Localization gap = import dependence x (6 - India readiness), range 1-25; Critical at 15 or more.
- Risk score = likelihood x impact; top three get mitigation actions with owner, timing and KPI; residual score shows expected reduction.

## Evidence and limits
Facts come from the press and government reports on the `Sources` sheet (cited as [n]). The 1-5 scores are my interpretation.
Three value-chain stages have no direct public source and are marked "Analyst estimate": verify before relying on them.
Plant status changes quickly; re-check before using any single item.

## Run
```bash
pip install openpyxl reportlab matplotlib
python build_p2.py        # then recalculate in Excel or LibreOffice
python build_p2_plan.py
```
