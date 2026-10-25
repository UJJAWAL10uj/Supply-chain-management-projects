"""AI sourcing copilot: supplier-quote comparison + disruption radar. Run: streamlit run app.py"""
import glob
import os

import altair as alt
import pandas as pd
import streamlit as st

from copilot import classify, compare, config, exposure, extract, llm, news

st.set_page_config(page_title="Sourcing Copilot", layout="wide")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

FIELDS = ["supplier_name", "country", "currency", "unit_price", "incoterm", "moq", "lead_time_days",
          "payment_terms_days", "advance_pct", "defect_rate_pct", "tooling_cost"]
NUM = ["unit_price", "moq", "lead_time_days", "payment_terms_days", "advance_pct", "defect_rate_pct", "tooling_cost"]
ARCH = {"India": "P-IN-TN", "China": "P-CN-SZ", "Vietnam": "P-VN-HCM"}
FLAG_COLOR = {"High": "#F8CBAD", "Medium": "#FFE699", "Low": "#C6EFCE"}

st.title("Sourcing Copilot")
st.caption("Compare supplier quotes on total cost, lead time, volume and quality, and see which mapped suppliers are exposed to "
           "live disruption signals. Sample data is synthetic or paraphrased; verify before acting.")

# ------------------------------------------------------------------ sidebar
client_ok = llm.available()
with st.sidebar:
    st.header("Engine")
    mode_label = st.radio("Extraction and classification", ["Rules only (offline)", "LLM-assisted with rules cross-check"],
                          index=0, help="LLM mode needs ANTHROPIC_API_KEY. Values the model cannot ground in the source text are dropped.")
    mode = "rules" if mode_label.startswith("Rules") else "both"
    if mode == "both" and not client_ok:
        st.warning("No ANTHROPIC_API_KEY found: running in rules mode.")
    st.header("Buying context")
    p = compare.Params()
    p.annual_volume = st.number_input("Annual volume (units)", 10_000, 50_000_000, p.annual_volume, step=50_000)
    p.order_qty = st.number_input("Typical order quantity", 1_000, 5_000_000, p.order_qty, step=1_000)
    p.contract_years = st.slider("Tooling amortisation (years)", 1, 5, p.contract_years)
    p.wacc = st.slider("Cost of capital", 0.0, 0.25, p.wacc, 0.01)
    p.cpq_multiple = st.slider("Cost of a shipped defect (x unit price)", 0.0, 10.0, p.cpq_multiple, 0.5)
    st.header("Decision weights")
    w = {k: st.slider(lbl, 0, 100, int(v * 100)) for (k, v), lbl in zip(p.weights.items(), ["Total cost", "Lead time", "MOQ / volume flexibility", "Quality"])}
    p.weights = {k: v / 100 for k, v in w.items()}
    p.risk_weight = st.slider("Weight of live disruption exposure", 0.0, 0.6, p.risk_weight, 0.05)
    with st.expander("FX to USD (edit before use)"):
        fx_df = st.data_editor(pd.DataFrame({"currency": list(p.fx), "to_usd": list(p.fx.values())}), hide_index=True, key="fx")
        p.fx = dict(zip(fx_df["currency"], fx_df["to_usd"]))

tab1, tab2, tab3 = st.tabs(["1. Quote comparison", "2. Disruption radar", "3. Risk-adjusted decision"])


def to_df(quotes):
    rows = []
    for q in quotes:
        r = {k: q.get(k) for k in FIELDS}
        r["flags"] = "; ".join(q.get("flags", []))
        r["source"] = q.get("source_name", "")
        rows.append(r)
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ tab 1
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        use_sample = st.checkbox("Use the 6 sample quotes (synthetic)", value=True)
        up = st.file_uploader("Or upload quote text files (.txt)", type=["txt"], accept_multiple_files=True)
    with c2:
        pasted = st.text_area("Or paste one quote", height=130, placeholder="Paste supplier quotation text here...")
    texts = []
    if use_sample:
        for f in sorted(glob.glob("data/quotes/q*.txt")):
            texts.append((os.path.basename(f), open(f).read()))
    for f in up or []:
        texts.append((f.name, f.read().decode("utf-8", "replace")))
    if pasted.strip():
        texts.append(("pasted_quote", pasted))
    sig = (tuple((n, hash(t)) for n, t in texts), mode)
    if st.session_state.get("quotes_sig") != sig:
        st.session_state["quotes"] = [extract.extract(t, n, mode=mode, client=llm.get_client() if mode == "both" else None) for n, t in texts]
        st.session_state["quotes_sig"] = sig
    quotes = st.session_state["quotes"]
    if not quotes:
        st.info("Add at least one quote to compare.")
    else:
        st.subheader("Extracted terms (edit any cell to correct the parse)")
        edited = st.data_editor(to_df(quotes), hide_index=True, width="stretch", key=f"ed_{hash(sig)}",
                                disabled=["flags", "source"])
        for q, (_, row) in zip(quotes, edited.iterrows()):
            for k in FIELDS:
                v = row[k]
                q[k] = None if pd.isna(v) else (float(v) if k in NUM else v)
        with st.expander("Evidence and flags per quote"):
            for q in quotes:
                st.markdown(f"**{q.get('supplier_name') or q.get('source_name')}** ({q['mode']} mode)")
                st.json({"evidence": q.get("evidence", {}), "flags": q.get("flags", [])}, expanded=False)

        df = compare.evaluate(quotes, p)  # cost / speed / volume / quality only; live risk is added in tab 3
        if df.empty:
            st.warning("No quote has a usable unit price.")
        else:
            st.subheader("Comparison")
            show = df[["rank", "supplier", "country", "price_usd", "tco", "lead_days", "moq", "terms_days", "advance_pct", "defect_pct", "decision_score", "assumed", "flags"]]
            st.dataframe(show.style.format({"price_usd": "${:.2f}", "tco": "${:.2f}", "decision_score": "{:.1f}"}), hide_index=True, width="stretch")
            comp = ["price_usd", "landing", "financing", "safety_stock", "cycle_stock", "quality_cost", "tooling_unit"]
            long = df.melt(id_vars=["supplier"], value_vars=comp, var_name="component", value_name="usd")
            cc1, cc2 = st.columns(2)
            with cc1:
                st.altair_chart(alt.Chart(long).mark_bar().encode(
                    x=alt.X("supplier:N", sort=list(df["supplier"]), axis=alt.Axis(labelAngle=-30, title=None)),
                    y=alt.Y("usd:Q", title="USD per unit"), color="component:N", tooltip=["supplier", "component", alt.Tooltip("usd:Q", format=".3f")]
                ).properties(title="Total cost per unit, by component", height=320), width="stretch")
            with cc2:
                st.altair_chart(alt.Chart(df).mark_circle(opacity=0.8).encode(
                    x=alt.X("lead_days:Q", title="Lead time (days)"), y=alt.Y("tco:Q", title="Total cost per unit (USD)", scale=alt.Scale(zero=False)),
                    size=alt.Size("moq:Q", title="MOQ"), color=alt.Color("defect_pct:Q", title="Defect %", scale=alt.Scale(scheme="redyellowgreen", reverse=True)),
                    tooltip=["supplier", "lead_days", alt.Tooltip("tco:Q", format=".2f"), "moq", "defect_pct"]
                ).properties(title="Cost / speed / volume / quality trade-off", height=320), width="stretch")
            st.subheader("Negotiation points")
            for sup, pts in compare.negotiation_points(df, quotes, p).items():
                with st.expander(sup):
                    for t in pts:
                        st.markdown(f"- {t}")
            st.download_button("Download comparison (CSV)", df.to_csv(index=False), "quote_comparison.csv", "text/csv")

# ------------------------------------------------------------------ tab 2
with tab2:
    src = st.radio("News source", ["Sample snapshot (offline, Oct 2026)", "Live RSS (needs internet)", "Paste text"], horizontal=True)
    items = []
    if src.startswith("Sample"):
        items = news.load_snapshot()
    elif src.startswith("Live"):
        qs = st.text_area("Search queries (one per line)", "\n".join(news.DEFAULT_QUERIES), height=120).splitlines()
        if st.button("Fetch live news"):
            got = []
            for i, q in enumerate([q for q in qs if q.strip()]):
                got += news.fetch_rss(q, prefix=f"R{i}-")
            st.session_state["live_items"] = news.dedupe(got)
        items = st.session_state.get("live_items", [])
        if not items:
            st.info("No live items yet (or no network). Click 'Fetch live news', or switch to the sample snapshot.")
    else:
        txt = st.text_area("Paste headlines or articles (blank line between items; first line is the title)", height=160)
        items = news.from_pasted(txt) if txt.strip() else []
    if items:
        client = llm.get_client() if mode == "both" else None
        events = [classify.classify(i, mode=mode, client=client) for i in items]
        reg = exposure.load_registry()
        radar = exposure.score(events, reg)
        st.session_state["radar"], st.session_state["events"] = radar, events
        st.caption(f"Scoring as of {radar['as_of']} with a {config.RECENCY_HALF_LIFE_DAYS}-day recency half-life. Flags: Medium at {config.FLAG_MEDIUM:.0f}, High at {config.FLAG_HIGH:.0f}.")
        ev_df = pd.DataFrame([dict(id=e["id"], date=e["date"], title=e["title"], type=e["primary_type"], direction=e["direction"], severity=e["severity"],
                                   places=", ".join(sorted(set().union(*[set(v) for v in e["regions"].values()]))), confidence=e["confidence"], engine=e["mode"], url=e["url"])
                              for e in events])
        min_sev = st.slider("Minimum severity shown", 0, 5, 0)
        st.subheader("Classified events")
        st.dataframe(ev_df[ev_df["severity"] >= min_sev], hide_index=True, width="stretch",
                     column_config={"url": st.column_config.LinkColumn("link")})
        st.subheader("Supplier exposure")
        t = radar["table"]
        st.dataframe(t.style.map(lambda v: f"background-color: {FLAG_COLOR.get(v, '')}", subset=["flag"]).format({"exposure": "{:.1f}", "priority": "{:.1f}"}),
                     hide_index=True, width="stretch")
        if not radar["heat"].empty:
            hm = radar["heat"]
            st.altair_chart(alt.Chart(hm).mark_rect().encode(
                x=alt.X("event_type:N", title=None, axis=alt.Axis(labelAngle=-30)), y=alt.Y("supplier:N", title=None, sort=list(t["supplier"])),
                color=alt.Color("contribution:Q", scale=alt.Scale(scheme="orangered"), title="Contribution"),
                tooltip=["supplier", "event_type", "contribution"]).properties(height=380, title="Which event types hit which supplier"), width="stretch")
        st.subheader("Why a supplier is flagged")
        pick = st.selectbox("Supplier", list(t["supplier"]))
        sid = t.loc[t["supplier"] == pick, "id"].iloc[0]
        det = sorted(radar["details"][sid], key=lambda d: -d["contribution"])
        if det:
            st.dataframe(pd.DataFrame(det)[["event_id", "title", "type", "severity", "weight", "recency", "contribution", "reason", "url"]],
                         hide_index=True, width="stretch", column_config={"url": st.column_config.LinkColumn("link")})
            st.markdown("**Suggested actions**")
            for typ in dict.fromkeys(d["type"] for d in det[:4]):
                st.markdown(f"- *{typ}*: {classify.ACTIONS[typ]}")
        else:
            st.write("No risk events matched this supplier.")
        with st.expander("Easing and opportunity signals"):
            st.write("Easing:", [f"{e['supplier']}: {e['title']}" for e in radar["easing"][:12]])
            st.write("Opportunities:", {s: [d["title"] for d in v] for s, v in radar["opps"].items() if v})
        st.subheader("Briefing")
        flagged = t[t["flag"] == "High"]
        facts = "\n".join(f"- {r.supplier}: exposure {r.exposure}, driver: {r.top_driver}" for r in flagged.itertuples())
        brief = None
        if mode == "both" and client is not None:
            brief = llm.ask_text("You write concise sourcing risk briefings. Use only the facts given.", f"Write a 5-bullet briefing from:\n{facts}", client=client)
        st.markdown(brief or ("**High-exposure suppliers**\n" + (facts or "- none")))
    else:
        st.session_state.pop("radar", None)

# ------------------------------------------------------------------ tab 3
with tab3:
    radar = st.session_state.get("radar")
    quotes = st.session_state.get("quotes", [])
    if not quotes or radar is None:
        st.info("Load quotes (tab 1) and news (tab 2) first.")
    else:
        st.write("Each quote supplier is matched to a risk archetype by country, so live disruption signals reduce its score. "
                 "Adjust the weight of live exposure in the sidebar (0 turns it off).")
        tbl = radar["table"].set_index("id")["exposure"].to_dict()
        expo = {(q.get("supplier_name") or q.get("source_name")): tbl.get(ARCH.get(q.get("country"), ""), 0.0) for q in quotes}
        base = compare.evaluate(quotes, p)
        adj = compare.evaluate(quotes, p, exposure=expo)
        if not adj.empty:
            m = base[["supplier", "rank", "decision_score"]].rename(columns={"rank": "rank_before", "decision_score": "score_before"}).merge(
                adj[["supplier", "country", "risk_exposure", "adjusted_score", "rank"]].rename(columns={"rank": "rank_after"}), on="supplier")
            m["move"] = m["rank_before"] - m["rank_after"]
            st.dataframe(m.sort_values("rank_after").style.format({"score_before": "{:.1f}", "adjusted_score": "{:.1f}", "risk_exposure": "{:.0f}"}),
                         hide_index=True, width="stretch")
            top = m.sort_values("rank_after").iloc[0]
            st.success(f"Top after risk adjustment: **{top.supplier}** (score {top.adjusted_score:.1f}, exposure {top.risk_exposure:.0f}).")
            ms = m.sort_values("rank_after")
            lines = ["| supplier | rank before | rank after | exposure | score |", "|---|---|---|---|---|"]
            lines += [f"| {r.supplier} | {r.rank_before} | {r.rank_after} | {r.risk_exposure:.0f} | {r.adjusted_score:.1f} |" for r in ms.itertuples()]
            md = "# Sourcing recommendation\n\n" + "\n".join(lines) + f"\n\nAs of {radar['as_of']}; exposure weight {p.risk_weight}."
            st.download_button("Download recommendation (Markdown)", md, "recommendation.md")
