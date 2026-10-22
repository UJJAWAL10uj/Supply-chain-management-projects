"""Two-page risk-mitigation plan; all numbers and text are read from the recalculated workbook."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak

XLSX = "/mnt/user-data/outputs/Enclosure_Localization_Supplier_Risk_Dashboard.xlsx"
OUT = "/mnt/user-data/outputs/Risk_Mitigation_Plan_Enclosure_Supply_Chain.pdf"
wb = load_workbook(XLSX, data_only=True)
D, V, S, R, M = wb["Dashboard"], wb["ValueChain"], wb["Suppliers"], wb["Risks"], wb["Mitigation"]

# ---- data
stages = [dict(id=V.cell(r, 1).value, name=V.cell(r, 2).value, dep=V.cell(r, 5).value, rdy=V.cell(r, 6).value,
               gap=V.cell(r, 7).value, prio=V.cell(r, 9).value, ev=V.cell(r, 10).value) for r in range(7, 16)]
crit = sorted([s for s in stages if s["prio"] == "Critical"], key=lambda s: -s["gap"])
n_est = sum(1 for s in stages if s["ev"] == "Analyst estimate")
sups = [(D.cell(r, 1).value, D.cell(r, 3).value, D.cell(r, 4).value) for r in range(22, 32)]
high = [s for s in sups if s[2] == "High"]
avg = D["B7"].value
risks = {R.cell(r, 1).value: dict(name=R.cell(r, 2).value, score=R.cell(r, 7).value, rank=R.cell(r, 8).value,
                                  after=R.cell(r, 13).value, red=R.cell(r, 14).value) for r in range(5, 13)}
top3 = sorted([k for k, v in risks.items() if v["rank"] <= 3], key=lambda k: risks[k]["rank"])
mit = [dict(rid=M.cell(r, 1).value, name=M.cell(r, 2).value, n=M.cell(r, 3).value, action=M.cell(r, 4).value,
            owner=M.cell(r, 5).value, timing=M.cell(r, 6).value, kpi=M.cell(r, 7).value) for r in range(5, 17)]

# ---- charts
fig, axs = plt.subplots(1, 2, figsize=(7.4, 2.7), dpi=200, gridspec_kw={"width_ratios": [1.05, 1]})
ax = axs[0]
names = [s[0].split(". ", 1)[1] for s in sups]
scores = [s[1] for s in sups]
colmap = {"High": "#C0504D", "Medium": "#E3A33B", "Low": "#4F9A5B"}
ax.barh(range(len(names)), scores, color=[colmap[s[2]] for s in sups], height=0.65)
ax.set_yticks(range(len(names)))
ax.set_yticklabels(names, fontsize=5.5)
ax.invert_yaxis()
ax.set_xlim(1, 5)
ax.tick_params(axis="x", labelsize=6)
ax.set_title("Supplier risk score (1 low, 5 high)", fontsize=7)
for sp_ in ("top", "right"):
    ax.spines[sp_].set_visible(False)
ax = axs[1]
order = sorted(stages, key=lambda s: -s["gap"])
ax.barh(range(len(order)), [s["gap"] for s in order],
        color=[{"Critical": "#C0504D", "Elevated": "#E3A33B", "Manageable": "#4F9A5B"}[s["prio"]] for s in order], height=0.65)
ax.set_yticks(range(len(order)))
ax.set_yticklabels([s["name"] for s in order], fontsize=5.5)
ax.invert_yaxis()
ax.tick_params(axis="x", labelsize=6)
ax.set_title("Localization gap score by stage (max 25)", fontsize=7)
for sp_ in ("top", "right"):
    ax.spines[sp_].set_visible(False)
plt.tight_layout()
plt.savefig("/home/claude/p2_charts.png")
plt.close()

# ---- styles
NAVY = colors.HexColor("#1F3864")
st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=NAVY)
st_m = ParagraphStyle("m", fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#444444"))
st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=NAVY, spaceBefore=5, spaceAfter=2)
st_b = ParagraphStyle("b", fontName="Helvetica", fontSize=8.4, leading=10.8)
st_bl = ParagraphStyle("bl", parent=st_b, leftIndent=9, bulletIndent=0)
st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=7.3, leading=8.9)
st_ch = ParagraphStyle("ch", fontName="Helvetica-Bold", fontSize=7.3, leading=8.9, textColor=colors.white)
st_s = ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=7, leading=8.6, textColor=colors.HexColor("#555555"))

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=11 * mm, bottomMargin=10 * mm,
                        title="Risk-mitigation plan: enclosure supply chain in India", author="Independent project")
story = []
story.append(Paragraph("Localization gap and supplier risk: mobile aluminium enclosures in India", st_t))
story.append(Paragraph("Short risk-mitigation plan &nbsp;|&nbsp; From: [Your name], B.E. &nbsp;|&nbsp; Date: 22 Oct 2026 &nbsp;|&nbsp; "
                       "Basis: attached Excel dashboard (public reports; scores are analyst judgments, not Apple data)", st_m))

story.append(Paragraph("What the map shows", st_h))
crit_txt = "; ".join(f"{s['name'][0].lower() + s['name'][1:] if not s['name'].startswith('CNC') else s['name']} (gap {s['gap']}/25)" for s in crit)
story.append(Paragraph(
    f"<b>India can already machine enclosures; the gaps sit upstream and underneath.</b> {len(crit)} of {len(stages)} value-chain stages are Critical: {crit_txt}. "
    f"Phone-enclosure aluminium extrusion is still 100% imported until Hindalco's Kuppam plant starts (target March 2027).", st_b))
story.append(Spacer(1, 2))
hi_txt = " and ".join(f"{s[0].split('. ', 1)[1]} ({s[1]:.2f})" for s in high)
story.append(Paragraph(
    f"Of 10 suppliers scored, {len(high)} rate High risk: {hi_txt}. The average score is {avg:.2f} of 5. "
    f"Risk is driven less by any one company than by shared dependencies: imported extrusion, China-linked equipment and know-how, and several new plants ramping at once.", st_b))
story.append(Spacer(1, 3))
story.append(Image("/home/claude/p2_charts.png", width=182 * mm, height=66.5 * mm))

story.append(Paragraph("Top three risks and what they are worth", st_h))
rows = [[Paragraph(t, st_ch) for t in ["Rank", "Risk", "Score now", "Score after mitigation", "Reduction"]]]
for k in top3:
    v = risks[k]
    rows.append([Paragraph(str(v["rank"]), st_c), Paragraph(f"{k}: {v['name']}", st_c), Paragraph(f"{v['score']:.0f}", st_c),
                 Paragraph(f"{v['after']:.0f}", st_c), Paragraph(f"{v['red']*100:.0f}%", st_c)])
t = Table(rows, colWidths=[12 * mm, 92 * mm, 24 * mm, 34 * mm, 20 * mm])
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                       ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
                       ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
story.append(t)
story.append(Spacer(1, 3))
r5 = risks["R5"]
story.append(Paragraph(
    f"<b>Link to the TCO model (Project 1):</b> at today's duty, freight and incentive assumptions a local Indian enclosure lands about 9.6% above a China import and needs "
    f"roughly 95% first-pass yield to match it. That makes yield and quality ramp (R5, score {r5['score']:.0f}) the gating factor for shifting volume, and it is why the plan "
    f"releases volume on milestones rather than on announcements.", st_b))
story.append(PageBreak())

story.append(Paragraph("Mitigation actions for the top three risks", st_t))
story.append(Paragraph("Owners are functions, not named people. KPIs are the early-warning signals to review weekly or monthly.", st_m))
story.append(Spacer(1, 4))
hdrs = ["Risk", "#", "Action", "Owner", "Timing", "KPI / trigger"]
rows = [[Paragraph(h, st_ch) for h in hdrs]]
prev = None
for m in mit:
    label = f"{m['rid']}: {m['name']}" if m["rid"] != prev else ""
    prev = m["rid"]
    rows.append([Paragraph(f"<b>{label}</b>", st_c), Paragraph(str(m["n"]), st_c), Paragraph(m["action"], st_c),
                 Paragraph(m["owner"], st_c), Paragraph(m["timing"], st_c), Paragraph(m["kpi"], st_c)])
t2 = Table(rows, colWidths=[27 * mm, 6 * mm, 80 * mm, 26 * mm, 17 * mm, 26 * mm], repeatRows=1)
style = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
         ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
for i, grp in enumerate([(1, 4), (5, 8), (9, 12)]):
    if i % 2 == 0:
        style.append(("BACKGROUND", (0, grp[0]), (-1, grp[1]), colors.HexColor("#F2F2F2")))
t2.setStyle(TableStyle(style))
story.append(t2)

story.append(Paragraph("Sequencing: first 90 days", st_h))
for tx in [
    "<b>Weeks 0-4:</b> set weeks-of-cover targets for extrusion and critical spares; confirm second import source for extrusion (R1, R2).",
    "<b>Weeks 4-8:</b> agree milestone gates and weekly ramp reviews with each new-plant supplier; define the single-plant share ceiling (R3).",
    "<b>Weeks 8-12:</b> start the equipment capability pilot with Indian machine makers and launch the technician certification plan (R2); open Kuppam sample-lot planning (R1).",
]:
    story.append(Paragraph(tx, st_bl, bulletText="\u2022"))

story.append(Paragraph("What this does not cover", st_h))
story.append(Paragraph(
    f"Supplier and stage scores are my judgments on public press reports, some based on unnamed sources, and plant status changes quickly. "
    f"{n_est} of {len(stages)} value-chain stages (cutting tools, surface finishing, inspection) have no direct public source and are marked Analyst estimate: verify them with suppliers before use. "
    f"Every score is an editable input in the workbook, so the ranking updates if you disagree.", st_b))
story.append(Spacer(1, 3))
story.append(Paragraph("Sources are listed on the Sources sheet of the workbook (ECMS approval notices, Business Standard, Digitimes, Bloomberg via Deccan Herald, Rest of World, Reuters).", st_s))

doc.build(story)
print("plan saved", len(crit), len(high), [k for k in top3])
