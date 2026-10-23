# Manufacturing Strategy, Supply Chain Risk & AI Sourcing Portfolio

A four-project portfolio exploring manufacturing cost competitiveness,
localization strategy, production-system improvement, and AI-assisted
procurement. The projects are linked through a common illustrative
product context: **mobile aluminium enclosures**.

**Project status:** Projects 1--3 completed; Project 4 in progress.\
**Data policy:** This is an independent project portfolio using
public-source research, analyst assumptions, and simulated production
data where stated. It does not use or imply access to confidential
company data.

------------------------------------------------------------------------

## Contents

1.  [India vs. China vs. Vietnam --- TCO and Scenario
    Model](#1-india-vs-china-vs-vietnam--tco-and-scenario-model)
2.  [Localization Gap and Supplier Risk
    Map](#2-localization-gap-and-supplier-risk-map)
3.  [Debottlenecking Study --- Theory of
    Constraints](#3-debottlenecking-study--theory-of-constraints)
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
approach. The aim is to look beyond quoted unit price and make cost
trade-offs visible under different operating scenarios.

### Model scope

The project brief defines a component-level model that considers
relevant cost drivers such as:

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

### Intended deliverable

An Excel-based cost and scenario model, supported by a one-page decision
memo that explains the recommendation, key assumptions, and sensitivity
to changing conditions.

> **Important:** The figures and comparisons are model outputs, not
> universal country-level cost claims. Interpret them in light of the
> model's input assumptions and verify current tariffs, freight,
> incentives, and supplier quotations before making a real sourcing
> decision.

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

The project includes an Excel dashboard that:

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
Plant timelines and supplier conditions can change; re-check the cited
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
-   `p3_code/simulate.py` --- simulation logic
-   `p3_code/run_experiments.py`- scenario experiments
-   `p3_code/build_p3.py` and `p3_code/build_p3_memo.py`-
    output-generation scripts
-   `p3_code/results.json`- experiment results
-   `README_project3.md`- project-specific documentation

### Limitations

Stage cycle times, intervention costs, and financial benefits are
assumptions. OEE ranges are benchmark-informed, but the model should be
recalibrated using real-time studies, maintenance records, production
targets, and supplier quotations before operational use. Financial gains
are demand-capped, and the model does not assign a dollar value to WIP
or lead-time improvements.

------------------------------------------------------------------------

## 4. AI Sourcing Copilot for Supplier Terms and Disruptions

**Status:** In progress\
**Tools planned/used in the partial implementation:** Python, optional
LLM API, Streamlit front end

### Objective

Build a small AI-assisted sourcing tool that combines supplier quote
comparison with supply-disruption monitoring. The intended workflow has
two parts:

1.  Extract key commercial terms from supplier quotations and compare
    suppliers.
2.  Monitor disruption signals and identify which mapped suppliers may
    be exposed.

### Intended capabilities

#### A. Supplier quote analysis

The intended input is a set of sample or synthetic supplier quotations.
The tool is designed to extract and structure:

-   Price
-   Minimum order quantity (MOQ)
-   Lead time
-   Payment terms

The comparison view is intended to help users examine cost, volume, and
quality trade-offs. Extracted terms should be checked against the
original quote before use.

#### B. Disruption monitoring and supplier exposure

The second component is intended to scan disruption news or feeds for
signals such as:

-   Port delays
-   Export controls
-   Tariff changes
-   Other supply interruptions relevant to mapped suppliers

The workflow reuses the supplier registry from Project 2, classifies
disruption events, and matches events to suppliers to estimate exposure.

### Implementation progress captured in the project notes

The work had started with:

-   A supplier registry based on the Project 2 dashboard
-   A news snapshot
-   A disruption classifier using a rule-based engine and optional LLM
    engine across a nine-type taxonomy
-   An exposure engine to match classified events to suppliers and score
    exposure

The shared project notes end while implementation is still in progress.
They do **not** confirm that the full application, supplier quote
extraction, Streamlit interface, end-to-end testing, or final
deliverables have been completed.

### Intended deliverable

A small Python application, potentially with a Streamlit interface, that
displays normalized supplier terms alongside relevant disruption events
and supplier exposure flags. The completed version should include source
links, timestamps, transparent matching logic, and a way to review or
override classifications.

### Safety and reliability considerations

AI-generated extractions and disruption classifications should be
treated as decision support, not authoritative conclusions. The tool
should preserve source evidence, show uncertainty, avoid presenting a
news mention as proof of supplier impact, and require human review
before commercial or sourcing action.

------------------------------------------------------------------------

## 5. How the Projects Connect

These projects form a connected decision-making workflow:

1.  **Compare country economics:** The TCO model establishes a framework
    for comparing India, China, and Vietnam under explicit cost and
    scenario assumptions.
2.  **Understand localization and supply risk:** The localization
    dashboard identifies capability gaps and supplier dependencies
    within the mobile-enclosure value chain.
3.  **Improve production performance:** The debottlenecking study
    examines how process constraints, changeovers, and release control
    affect throughput and flow.
4.  **Support ongoing sourcing decisions:** The AI sourcing copilot is
    being developed to organize supplier terms and surface disruption
    signals against the supplier map.

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

  AI sourcing       Which supplier     Term extraction,  Application in
  copilot           terms matter, and  event             progress
                    what disruption    classification,   
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
    ├── classify.py
    ├── exposure.py
    ├── supplier_registry/
    ├── data_samples/
    └── app.py
```

This is a suggested organization, not a claim that every file or folder
already exists. Keep only the files actually present in the repository
and update the project-specific README files with exact setup
instructions.

------------------------------------------------------------------------

## 7. Interpretation and Limitations

-   **Independent work:** These projects are portfolio analyses, not
    official company studies. Do not imply access to Apple or any other
    company's confidential sourcing, supplier, or manufacturing data.
-   **Public-source research:** Supplier and localization claims depend
    on the sources and reporting available when the research was
    performed. Re-check time-sensitive facts before interviews or
    decisions.
-   **Model assumptions:** Clearly distinguish observed facts,
    public-source findings, analyst judgments, and simulated inputs.
-   **No invented performance claims:** Report numerical outcomes only
    with their assumptions and context. The debottlenecking results are
    simulation outputs, not measured factory improvements.
-   **Human review:** Supplier ratings, risk scores, AI-extracted
    commercial terms, and disruption matches require validation before
    real-world use.
-   **Reproducibility:** Update machine-specific file paths and
    configuration before running scripts in another environment. Do not
    commit API keys, credentials, or confidential supplier documents.

## Summary

The portfolio combines **TCO and scenario analysis, localization
strategy, supplier risk assessment, Theory of Constraints,
discrete-event simulation, and AI-assisted sourcing**. Its central theme
is turning cost, supply-chain, and operational information into
structured decisions---while making assumptions, evidence, and
uncertainty explicit.
