Sourcing Copilot

A small tool that does two jobs for a sourcing team:

Quote comparison. Reads free-text supplier quotes, extracts price, MOQ, lead time, payment terms, Incoterm, quality and tooling, and compares them on a like-for-like total cost per unit plus lead-time, volume and quality trade-offs, with number-backed negotiation points.
Disruption radar. Classifies news (port delays, export controls, tariffs, freight, plant incidents, commodity prices...) and flags which mapped suppliers are exposed, with the reason, the recency-weighted score and a suggested action.

A third tab combines both: live exposure reduces a supplier's decision score, so you can see whether risk changes the ranking.

SYNTHETIC / PARAPHRASED SAMPLE DATA. Quotes are invented. News items are my paraphrases of public reports (links in data/news_snapshot.json). Registry rows come from the Project 2 dashboard (public reports). FX and landing-cost adders are editable defaults.

Run
bash
pip install -r requirements.txt
streamlit run app.py            # rules mode works offline
export ANTHROPIC_API_KEY=...    # optional: enables LLM mode (COPILOT_MODEL overrides the model name)
pytest -q                       # 13 tests, no network needed
python evaluate.py rules        # accuracy against hand labels (python evaluate.py llm needs a key)
How it works
quote text --> extract (rules | LLM + evidence) --> ground + cross-check --> editable table --> TCO + scores --> ranking, charts, negotiation
news items --> classify (rules | LLM) --> event {type, direction, severity, places} --> exposure vs registry --> ranked suppliers, heat map, actions
Total cost per unit = price + landing add-on by Incoterm + financing (advance vs credit days) + safety stock + cycle stock (MOQ)
cost of poor quality + tooling per unit. Landing add-ons are calibrated to the Project 1 TCO model (China FOB import = about 17% of price).
Guards against silent errors. LLM values must be backed by an evidence snippet that literally appears in the quote, or they are dropped. When LLM and rules disagree, the field is flagged. Prices quoted per 1,000 are normalised and flagged. Prices far from peers are flagged. Missing fields are filled with flagged assumptions, never silently. Every extracted value is editable in the UI before scoring.
Exposure per event and supplier = severity/5 x relevance weight x recency (45-day half-life). Weights: direct mention 1.0, lane 0.8, site city 0.7, gateway port 0.55, state 0.5, China-linked or aluminium dependency 0.6 / 0.5 (partial = half), parent country 0.4, sea-freight 0.25, country-wide 0.2. Per supplier, the largest contribution per event type is kept, then combined across types. Only risk events add exposure; easing and opportunity events are listed separately.
Evaluation (see evaluate.py, results in docs/)

Labels were written by the author, so this is a smoke test, not a benchmark. Rules were written and tuned on the "dev" items only; "test" and "hard" items were scored without tuning (exception: a per-1,000 price normaliser was added after the hard set exposed a silent unit error). Rules mode: quote fields 100% on dev and test (33/33 each), 72% on the harder set (38/53, with 2 silent wrong numbers); news classification on 19 labelled items (17 events plus 2 noise items): primary type among predicted 100% (17/17), exact primary type 94% (16/17), direction 100%, severity within one point 100%, noise rejected 2/2. The hard set shows where rules break (amount before currency, one-line emails, "45 days credit", summed lead-time phrasing): that is the case for LLM mode with grounding. LLM mode is implemented and unit-tested with a fake client but was NOT run against the live API in the build environment (no key); run python evaluate.py llm with your key to measure it.

Limits

Small, author-labelled test sets. Live RSS fetch is implemented (Google News search feeds) and its parser is tested, but it was not run live in the build environment. Gateway ports, dependency flags and archetype risk profiles are analyst assumptions. Exposure is a screening score, not a forecast. Always read the cited article before acting.
