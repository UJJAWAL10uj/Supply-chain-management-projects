# Manufacturing Strategy, Supply Chain Risk & AI Sourcing Portfolio

A four-project portfolio exploring manufacturing cost competitiveness,
localization strategy, production-system improvement, and AI-assisted
procurement. The projects are connected through a common illustrative
product context: **mobile aluminium enclosures**.

**Portfolio status: All four projects completed**\
**Data policy:** Independent portfolio work using public-source
research, analyst assumptions, and simulated/sample data where stated.
This portfolio does not imply access to confidential company data.

## Contents

1.  [India vs. China vs. Vietnam- TCO and Scenario
    Model](#1-india-vs-china-vs-vietnam-tco-and-scenario-model)
2.  [Localization Gap and Supplier Risk
    Map](#2-localization-gap-and-supplier-risk-map)
3.  [Debottlenecking Study- Theory of
    Constraints](#3-debottlenecking-study-theory-of-constraints)
4.  [AI Sourcing Copilot for Supplier Terms and
    Disruptions](#4-ai-sourcing-copilot-for-supplier-terms-and-disruptions)
5.  [How the Projects Connect](#5-how-the-projects-connect)
6.  [Repository Structure](#6-repository-structure)
7.  [Interpretation and Limitations](#7-interpretation-and-limitations)

------------------------------------------------------------------------

## 1. India vs. China vs. Vietnam- TCO and Scenario Model

**Status:** Completed\
**Focus:** Country-level manufacturing and sourcing economics

### Objective

Compare manufacturing or sourcing a selected mechanical component across
India, China, and Vietnam using a **Total Cost of Ownership (TCO)**
approach. The model looks beyond quoted unit price to make cost
trade-offs visible under different operating scenarios.

### Model scope

The component-level model considers relevant cost drivers such as:

-   Material and machining costs
-   Yield loss and first-pass yield
-   Labour
-   Freight and duties
-   Applicable incentives, including relevant Indian PLI/ECMS schemes
    where eligible
-   Scenario changes such as tariff changes, freight spikes, and a ±10%
    volume swing
-   Sensitivity analysis to show which assumptions have the greatest
    effect on the comparison
-   Incoterm comparison: **FOB (Free on Board), CIF (Cost, Insurance and
    Freight), and DDP (Delivered Duty Paid)**, including responsibility
    for relevant costs

### Connection to the other projects

The localization analysis refers to the TCO model's indication that
India needs approximately **95% first-pass yield to match China** under
the model's assumptions. The debottlenecking study uses an **86%
first-pass yield** assumption to remain consistent with that project
context.

### Deliverable

An Excel-based cost and scenario model, supported by a one-page decision
memo that explains the recommendation, key assumptions, and sensitivity
to changing conditions.

> The figures and comparisons are model outputs, not universal
> country-level cost claims. Verify current tariffs, freight,
> incentives, exchange rates, and supplier quotations before making a
> real sourcing decision.

------------------------------------------------------------------------

## 2. Localization Gap and Supplier Risk Map

**Status:** Completed\
**Product category:** Mobile aluminium enclosures\
**Focus:** Domestic capability gaps, supplier dependency, and mitigation
planning

### Objective

Map the enclosure value chain to understand which capabilities are
available in India, where localization gaps remain, and which supplier
or ecosystem risks could affect production continuity.

### Method

The Excel dashboard:

-   Maps the mobile-enclosure value chain across **9 stages**
-   Scores **10 suppliers** against four criteria: single-source
    exposure, capacity, geopolitical exposure, and quality readiness
-   Uses editable scoring inputs, weights, and thresholds so rankings
    can be challenged and updated
-   Identifies the most important localization and supplier risks
-   Supports a risk-mitigation plan with owners, timing, and KPIs

### Main findings reported by the model

-   **Three of nine value-chain stages** were assessed as critical.
-   **Precision aluminium extrusion** was identified as a major
    localization gap, with the project research reporting 100% import
    dependence for the relevant extruded aluminium input.
-   **CNC machine tools/equipment** and **process know-how/skilled
    engineers** were also identified as critical gaps.
-   Among the 10 suppliers scored, **Yuzhan (Foxconn's India arm) scored
    4.00** and **Hindalco scored 3.75**, both rated High in the model;
    **Jabil India scored 2.25**, the lowest score in the assessment.
-   The top three risks were imported extrusion dependence (score 20),
    Chinese equipment and engineer dependence (score 20), and capacity
    ramp delays across new plants (score 16).
-   The mitigation plan contains **12 actions** across the top three
    risks, with owners, timing, and KPIs. Under the model, completing
    the actions reduces residual risk by an estimated **40--55%**.

### Deliverables

-   `Enclosure_Localization_Supplier_Risk_Dashboard.xlsx`- Excel
    dashboard and editable scoring model
-   `Risk_Mitigation_Plan_Enclosure_Supply_Chain.pdf`- two-page
    risk-mitigation plan
-   `README_project2.md`- project-specific documentation
-   Build scripts for generating the dashboard and plan

### Data and confidence notes

Supplier scores are analyst judgments based on public reporting, not
independently audited supplier ratings. Three value-chain
stages---cutting tools, surface finishing, and inspection---were marked
as analyst estimates because public sources were not available for them.
Some supplier facts came from press reports citing unnamed sources.
Plant timelines and supplier conditions can change; re-check cited
sources before relying on them.

------------------------------------------------------------------------

## 3. Debottlenecking Study- Theory of Constraints

**Status:** Completed\
**Product context:** One module of an illustrative mobile-enclosure
factory\
**Tools:** Python, SimPy, Excel\
**Focus:** Throughput, work in progress (WIP), flow time, and the
economics of operational fixes

### Objective

Model a multi-stage production line and apply the **Theory of
Constraints (TOC)** to identify the process limiting total output.
Simulate possible interventions and compare their throughput impact and
assumed cost.

### Model setup

The study models a **seven-stage enclosure production line** targeting
10,000 good units per week. The stages represent a manufacturing flow
that includes machining, finishing, inspection, and packing. The model
uses illustrative cycle times and intervention costs; it is a
simulation, not a study of a real factory.

The project includes a SimPy discrete-event simulation, an Excel model,
experiment results, and a two-page memo. The Excel and simulation
outputs were reported to reconcile closely: **8,932 units/week** in the
analytical model versus **8,938 units/week** in simulation for the
baseline.

### Findings reported by the model

-   **Anodising** was identified as the constraint, loaded to **112% of
    capacity** at target demand.
-   Baseline output was approximately **8,932 good units/week**, or
    about **89% of the 10,000-unit target**.
-   CNC OP20 finish was the next most heavily loaded stage, at
    approximately **91%**.
-   The best single intervention was **SMED** (reducing rack-changeover
    time), which reduced the assumed rack batch changeover from 52 to 46
    minutes and lifted simulated output to approximately **9,969
    units/week**.
-   Under assumed costs, SMED was modelled at a one-time cost of
    **\$40,000** and an annual net benefit of approximately **\$160,000
    per module**, with payback in under three months.
-   **Drum-buffer-rope (DBR)** release control had little effect on
    maximum output by itself, but substantially reduced WIP and flow
    time. The memo reports WIP falling by about **92%** and flow time
    decreasing from **67 hours to 5.3 hours**.
-   Adding a third anodising line was not economically attractive in the
    model because output gains were capped by demand; the estimated net
    result was approximately **−\$62,000/year**.
-   The combined recommendation was to implement SMED plus DBR, with a
    modelled net benefit of approximately **\$156,000/year per module**
    (or **\$1.56 million/year across ten modules**) under the stated
    assumptions, and to improve anodising maintenance headroom.
-   After the constraint is relieved, the model indicates that the
    bottleneck shifts to CNC OP20 finish.

### Important modelling lesson

A buffer sized for the original, slower constraint starved the line
after SMED increased the constraint's speed. Output fell to roughly
9,840 units until the buffer was resized from 400 to 900 units. This
illustrates why process changes should be evaluated at the system level,
not only at the workstation level.

### Deliverables

-   `Debottlenecking_Study_Enclosure_Line.xlsx`- Excel model and
    dashboard
-   `Debottlenecking_Memo_Enclosure_Line.pdf`- two-page findings and
    recommendation memo
-   `p3_code/simulate.py`- simulation logic
-   `p3_code/run_experiments.py`- scenario experiments
-   `p3_code/build_p3.py` and `p3_code/build_p3_memo.py`-
    output-generation scripts
-   `p3_code/results.json`- experiment results
-   `README_project3.md`- project-specific documentation

### Limitations

Stage cycle times, intervention costs, and financial benefits are
assumptions. OEE ranges are benchmark-informed, but the model should be
recalibrated using real time studies, maintenance records, production
targets, and supplier quotations before operational use. Financial gains
are demand-capped, and the model does not assign a dollar value to WIP
or lead-time improvements.

------------------------------------------------------------------------

## 4. AI Sourcing Copilot for Supplier Terms and Disruptions

**Status:** Completed- application built; LLM mode and live-news
fetching not yet verified in a live environment\
**Tools:** Python, Streamlit, rule-based extraction/classification,
optional LLM API

### Objective

Build an AI-assisted sourcing tool that combines supplier quote
comparison with supply-disruption monitoring. The copilot helps
structure commercial terms, compare supplier trade-offs, and flag mapped
suppliers that may be exposed to disruption signals.

### A. Supplier quote comparison

The app accepts free-text supplier quotes and extracts:

-   Price and volume price breaks
-   Minimum order quantity (MOQ)
-   Lead time
-   Payment terms
-   Incoterm
-   Quality information
-   Tooling costs

It displays supporting evidence for extracted values and allows users to
correct them. Supplier ranking considers total cost per unit, lead time,
volume flexibility, and quality, and produces numeric negotiation
points.

### B. Disruption radar

The app classifies news into **nine event types**, assigning a direction
(risk, easing, or opportunity) and severity. It scores supplier exposure
using recency-weighted signals and includes the reason, source link, and
a suggested action.

The registry includes the **10 suppliers from Project 2** plus **three
supplier archetypes**. A combined view allows disruption exposure to
adjust quote-supplier scores.

### Implementation and evaluation reported

-   Built as a **Streamlit application** with a two-page report.
-   **13 tests pass** in the reported test run.
-   Rules mode is designed to work offline and can be launched from the
    `sourcing_copilot/` directory using `streamlit run app.py`.
-   On sample quote data, quoted prices ranged from
    **\$3.18--\$4.14/unit**; after modelled landing cost, payment
    timing, MOQ, quality, and tooling adjustments, the comparable range
    was **\$3.85--\$4.39/unit**. An import quote ranked first under the
    sample assumptions, consistent with the Project 1 TCO context.
-   The disruption radar used **17 October 2026 signals** in the
    reported sample snapshot. Five registry entries were rated High
    exposure by the model: four Tamil Nadu/Karnataka enclosure-maker
    entries (Tata, Yuzhan, Motherson, and Aequs) and the India
    local-supplier archetype.
-   Adding disruption exposure did **not** change the quote ranking at
    any tested weight from 0 to 0.6; the report records this result
    rather than claiming an effect that was not observed.
-   Quote-field extraction was reported as **100% accurate on three
    development/test quotes** and **72% accurate (38/53 fields) on five
    deliberately messy quotes**. Two incorrect numeric values were
    returned instead of being marked missing. One failure involving a
    price quoted per 1,000 units was fixed.
-   News classification got the exact primary event type right on **16
    of 17 items**; direction, severity, and noise rejection were
    reported as correct on that test set.

### LLM and live-data verification status

The LLM integration is implemented but was tested only with a fake
client because no API key was available in the development environment.
A guard rejects values the model cannot support with a snippet from the
quote. Run `python evaluate.py llm` with the required API configuration
to evaluate the LLM mode.

Live news fetching was not verified in the reported environment due to
lack of network access, although the news parser was tested. The 13
passing tests and sample results should therefore be understood as
application/test-harness evidence, not proof that every external
integration has been validated live.

### Deliverables

-   `sourcing_copilot.zip`- application package
-   `Sourcing_Copilot_Report.pdf`- report
-   `sourcing_copilot/README.md`- setup and project documentation
-   `sourcing_copilot/app.py`- Streamlit application
-   The project notes also identify `classify.py`, `exposure.py`, and
    `evaluate.py` as relevant implementation/evaluation components.

------------------------------------------------------------------------

## 5. How the Projects Connect

The four projects form a connected decision workflow:

1.  **Compare country economics:** The TCO model establishes a framework
    for comparing India, China, and Vietnam under explicit cost and
    scenario assumptions.
2.  **Understand localization and supply risk:** The localization
    dashboard identifies capability gaps and supplier dependencies in
    the mobile-enclosure value chain.
3.  **Improve production performance:** The debottlenecking study
    examines how process constraints, changeovers, and release control
    affect throughput and flow.
4.  **Support ongoing sourcing decisions:** The sourcing copilot
    compares quote terms and overlays disruption signals against the
    supplier registry.

  ------------------------------------------------------------------------
  Project           Core question      Main method       Main output
  ----------------- ------------------ ----------------- -----------------
  TCO and scenario  Which sourcing     Cost modelling,   Excel model and
  model             location is        scenario and      recommendation
                    competitive under  sensitivity       memo
                    different          analysis          
                    assumptions?                         

  Localization and  Where are the      Value-chain       Excel dashboard
  supplier risk     capability gaps    mapping and       and mitigation
                    and supply         supplier risk     plan
                    vulnerabilities?   scoring           

  Debottlenecking   What limits        Theory of         Simulation, Excel
  study             output, and which  Constraints,      model, and memo
                    fix is worthwhile? SimPy simulation, 
                                       cost-benefit      
                                       analysis          

  AI sourcing       Which supplier     Quote extraction, Streamlit
  copilot           terms matter, and  event             application and
                    what disruption    classification,   report
                    signals require    supplier exposure 
                    review?            matching          
  ------------------------------------------------------------------------

------------------------------------------------------------------------

## 6. Repository Structure

A suggested top-level layout for storing the four projects together:

``` text
manufacturing-supply-chain-projects/
├── README.md
├── 01-tco-scenario-model/
│   ├── model/
│   ├── memo/
│   └── README.md
├── 02-localization-supplier-risk/
│   ├── Enclosure_Localization_Supplier_Risk_Dashboard.xlsx
│   ├── Risk_Mitigation_Plan_Enclosure_Supply_Chain.pdf
│   ├── README_project2.md
│   └── build_scripts/
├── 03-debottlenecking-toc/
│   ├── Debottlenecking_Study_Enclosure_Line.xlsx
│   ├── Debottlenecking_Memo_Enclosure_Line.pdf
│   ├── README_project3.md
│   └── p3_code/
└── 04-ai-sourcing-copilot/
    ├── README.md
    ├── app.py
    ├── classify.py
    ├── exposure.py
    ├── evaluate.py
    └── ...
```

This is a suggested organization, not a claim that every file or folder
already exists. Keep only files actually present in the repository and
update setup instructions to match the committed paths.

### Run Project 4 locally

From the `sourcing_copilot/` directory:

``` bash
streamlit run app.py
```

Rules mode is reported to work offline. LLM mode requires the
appropriate API configuration and should be evaluated separately. Do not
commit API keys or credentials.

------------------------------------------------------------------------

## 7. Interpretation and Limitations

-   **Independent work:** These are portfolio analyses, not official
    company studies. Do not imply access to Apple or another company's
    confidential sourcing, supplier, or manufacturing data.
-   **Public-source research:** Supplier and localization claims depend
    on the sources available when the research was performed. Re-check
    time-sensitive facts before interviews or decisions.
-   **Model assumptions:** Distinguish public-source findings, analyst
    judgments, synthetic inputs, and simulated outputs.
-   **No invented performance claims:** The debottlenecking results are
    simulation outputs, not measured factory improvements. The sourcing
    copilot's accuracy figures apply only to the stated sample sizes and
    tests.
-   **Live integration caveats:** Project 4's LLM mode and live news
    fetching were not verified end to end in the reported environment.
-   **Human review:** Supplier ratings, cost estimates, extracted quote
    terms, and disruption matches require validation before real-world
    use.
-   **Reproducibility:** Update machine-specific paths, currency rates,
    and configuration before running the project elsewhere.

## Summary

This portfolio combines **TCO and scenario analysis, localization
strategy, supplier risk assessment, Theory of Constraints,
discrete-event simulation, and AI-assisted sourcing**. Together, the
projects demonstrate an approach to turning cost, supply-chain,
production, and supplier information into structured decisions while
making assumptions, evidence, and uncertainty explicit.
