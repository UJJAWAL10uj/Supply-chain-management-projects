# TCO & Scenario Model: CNC Aluminium Enclosure, India vs China vs Vietnam

## Question
Where should a CNC-machined aluminium smartphone enclosure be sourced from to be delivered to a
final-assembly plant near Chennai: a local Indian supplier, or imports from China or Vietnam?
How robust is the answer to tariff, freight, volume, metal-price, incentive and yield shocks?

## Deliverables
| File | What it is |
|---|---|
| `Apple_Sourcing_TCO_Model_Aluminium_Enclosure.xlsx` | Formula-driven model (about 2,300 formulas, 0 errors): Inputs, 11 Scenarios, Calc engine, Summary, Sensitivity grids and break-evens, Incoterms (FOB/CIF/DDP), Sources |
| `Sourcing_Recommendation_Memo_Aluminium_Enclosure.pdf` | One-page recommendation memo (numbers read from the model) |
| `build_model.py` | Regenerates the workbook (openpyxl) |
| `build_memo.py` | Regenerates the memo from the recalculated workbook (reportlab, matplotlib) |

## Method
Landed cost per unit = net material (after chip-scrap credit) + machining + direct labour + finishing
+ yield loss + fixed cost per unit + supplier margin - ECMS incentive passed to buyer (India only)
+ freight + insurance + Basic Customs Duty + Social Welfare Surcharge + other duties + clearance and
haulage + pipeline carrying cost. IGST is excluded (creditable).

## Sourced vs assumed
- Sourced: LME aluminium price (6 Oct 2026), ECMS enclosure incentive range (6-8% initially), US/India tariff context.
- Assumed (illustrative, replace with RFQ or benchmark data): cycle times, machine and labour rates, yields,
  tooling, overhead, margins, freight rates, transit days. See the `Sources` sheet.
- Current Indian BCD for the enclosure's HS code; any preferential rate for Vietnam-origin parts.

## Run
```bash
pip install openpyxl reportlab matplotlib
python build_model.py   # then recalculate in Excel or LibreOffice so cached values exist
python build_memo.py
```
Just edit the blue cells on `Inputs` and `Scenarios`; everything will recalculate.
