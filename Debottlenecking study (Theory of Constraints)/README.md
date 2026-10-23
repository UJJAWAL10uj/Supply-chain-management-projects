# Debottlenecking an Aluminium-Enclosure Line (Theory of Constraints)

SIMULATED, ILLUSTRATIVE DATA. Not a real plant. Stage parameters are benchmark-informed assumptions; replace with time-study data.

## Question
One production module must deliver 10,000 good enclosures a week through seven stages (saw, CNC OP10, CNC OP20,
deburr, anodising, inspection, packing). Where is the constraint, how much output does it cost, which fixes recover it,
and what do they cost?

## Deliverables
| File | What it is |
|---|---|
| `Debottlenecking_Study_Enclosure_Line.xlsx` | Excel model (about 620 formulas, 0 errors): Inputs, BaseAnalysis (capacity, utilisation, takt, OEE, constraint), Fixes (12 scenarios, constraint migration), SimResults (SimPy output, rope sweeps), Economics (gain vs cost), Dashboard with 5 charts |
| `Debottlenecking_Memo_Enclosure_Line.pdf` | Two-page memo: finding, fixes compared, recommendation |
| `p3_code/` | `simulate.py` (SimPy model), `run_experiments.py` (sweeps and scenarios), `build_p3.py` (workbook), `build_p3_memo.py` (memo), `results.json` (simulation output) |

## Method
1. Identify: capacity = machines x planned minutes x availability x performance / ideal cycle; utilisation above 100% marks the constraint.
2. Exploit: SMED (rack batch 52 to 46 min), TPM, yield at the constraint.
3. Subordinate: drum-buffer-rope release gated by a buffer ahead of anodising.
4. Elevate: third anodising line, 7-day working, only if still needed.
5. Repeat: the constraint moves; the buffer must be re-sized.
SimPy adds breakdowns (operating-time MTBF / MTTR), cycle-time variability, stage scrap, 48-unit rack batching, WIP and lead time. 8 replications per scenario, common random numbers.

## Basis and limits
Benchmarks (world-class OEE about 85%, typical about 60%; machining world-class 80-85%) inform plausible ranges only. Fix costs are assumptions.
Price and material cost come from the Project 1 TCO model. Gains are capped at demand; WIP and lead-time benefits are not valued in dollars.

## Run
```bash
pip install simpy openpyxl reportlab matplotlib
cd p3_code
python run_experiments.py      # about 3-4 minutes, writes results.json
python build_p3.py             # then recalculate in Excel or LibreOffice
python build_p3_memo.py
```
(The scripts use the author's paths; edit the path constants at the top of each file.)
