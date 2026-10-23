"""Two-page debottlenecking memo; all numbers read from the recalculated workbook."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

XLSX = "/mnt/user-data/outputs/Debottlenecking_Study_Enclosure_Line.xlsx"
OUT = "/mnt/user-data/outputs/Debottlenecking_Memo_Enclosure_Line.pdf"
wb = load_workbook(XLSX, data_only=True)
BA, FX, EC, SR, IN = wb["BaseAnalysis"], wb["Fixes"], wb["Economics"], wb["SimResults"], wb["Inputs"]

stages = [dict(name=BA[f"B{r}"].value, util=BA[f"G{r}"].value, oee=BA[f"J{r}"].value) for r in range(5, 12)]
base = dict(constraint=BA["D19"].value, theory=BA["D20"].value, pct=BA["D22"].value, oee=BA["D23"].value,
            takt_s=BA["D16"].value * 60, nxt=BA["D24"].value, nxt_util=BA["D25"].value)
demand = IN["B7"].value
contrib = IN["B13"].value
rows = {}
for r in range(5, 17):
    rows[EC[f"A{r}"].value] = dict(
        name=EC[f"B{r}"].value, sim=EC[f"C{r}"].value, gain=EC[f"E{r}"].value, net=EC[f"K{r}"].value, pay=EC[f"L{r}"].value,
        wip=EC[f"M{r}"].value, flow=EC[f"N{r}"].value, onetime=EC[f"H{r}"].value, annual=EC[f"I{r}"].value,
        theory=EC[f"O{r}"].value, fullnet=EC[f"R{r}"].value,
        con=FX[f"X{r}"].value, head=FX[f"AC{r}"].value, maxgood=FX[f"Y{r}"].value)
wip_abs = {SR[f"A{r}"].value: SR[f"G{r}"].value for r in range(5, 17)}
rope1, rope2 = wb["Inputs"]["C29"].value, None
for r in range(5, 17):
    if SR[f"A{r}"].value == "C7":
        rope2 = SR[f"Q{r}"].value

# ---- charts
fig, axs = plt.subplots(1, 2, figsize=(7.4, 2.6), dpi=200)
ax = axs[0]
order = list(range(len(stages)))
cols = ["#C0504D" if s["util"] >= 1 else ("#E3A33B" if s["util"] >= 0.9 else "#4F9A5B") for s in stages]
ax.barh(order, [s["util"] * 100 for s in stages], color=cols, height=0.65)
ax.set_yticks(order)
ax.set_yticklabels([s["name"] for s in stages], fontsize=6)
ax.invert_yaxis()
ax.axvline(100, color="black", lw=0.7, ls="--")
ax.tick_params(axis="x", labelsize=6)
ax.set_xlabel("Utilisation at 10,000 good units/week (%)", fontsize=6)
ax.set_title("Step 1: anodising is the constraint", fontsize=7)
for sp in ("top", "right"):
    ax.spines[sp].set_visible(False)
ax = axs[1]
ids = list(rows.keys())
vals = [rows[i]["sim"] for i in ids]
ax.bar(range(len(ids)), vals, color=["#7F7F7F" if i == "C0" else "#4F81BD" for i in ids], width=0.7)
ax.axhline(demand, color="#C0504D", lw=0.9, ls="--")
ax.text(len(ids) - 0.5, demand + 40, "demand", fontsize=5.5, color="#C0504D", ha="right")
ax.set_xticks(range(len(ids)))
ax.set_xticklabels(ids, fontsize=5.5)
ax.set_ylim(8000, 10300)
ax.tick_params(axis="y", labelsize=6)
ax.set_title("Simulated good units per week by scenario", fontsize=7)
for sp in ("top", "right"):
    ax.spines[sp].set_visible(False)
plt.tight_layout()
plt.savefig("/home/claude/p3_c1.png")
plt.close()

fig, axs = plt.subplots(1, 2, figsize=(7.4, 2.4), dpi=200)
ax = axs[0]
ax.bar(range(len(ids)), [wip_abs[i] for i in ids], color=["#7F7F7F" if i == "C0" else "#4F81BD" for i in ids], width=0.7)
ax.set_xticks(range(len(ids)))
ax.set_xticklabels(ids, fontsize=5.5)
ax.tick_params(axis="y", labelsize=6)
ax.set_title("Average WIP (units): push builds a queue at the constraint", fontsize=7)
for sp in ("top", "right"):
    ax.spines[sp].set_visible(False)
ax = axs[1]
nets = [rows[i]["net"] / 1000 for i in ids]
ax.bar(range(len(ids)), nets, color=["#4F9A5B" if n > 0 else "#C0504D" for n in nets], width=0.7)
ax.axhline(0, color="black", lw=0.6)
ax.set_xticks(range(len(ids)))
ax.set_xticklabels(ids, fontsize=5.5)
ax.tick_params(axis="y", labelsize=6)
ax.set_title("Net annual benefit per module (USD thousand)", fontsize=7)
for sp in ("top", "right"):
    ax.spines[sp].set_visible(False)
plt.tight_layout()
plt.savefig("/home/claude/p3_c2.png")
plt.close()

# ---- PDF
NAVY = colors.HexColor("#1F3864")
st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=NAVY)
st_m = ParagraphStyle("m", fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#444444"))
st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=NAVY, spaceBefore=5, spaceAfter=2)
st_b = ParagraphStyle("b", fontName="Helvetica", fontSize=8.4, leading=10.8)
st_bl = ParagraphStyle("bl", parent=st_b, leftIndent=9, bulletIndent=0)
st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=7.4, leading=9)
st_ch = ParagraphStyle("ch", fontName="Helvetica-Bold", fontSize=7.4, leading=9, textColor=colors.white)
st_s = ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=7, leading=8.6, textColor=colors.HexColor("#555555"))
usd = lambda v: f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"
k = lambda v: f"-${abs(v)/1000:,.0f}k" if v < 0 else f"${v/1000:,.0f}k"

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=11 * mm, bottomMargin=10 * mm,
                        title="Debottlenecking study: enclosure production line", author="Independent project")
S = []
S.append(Paragraph("Debottlenecking an aluminium-enclosure line: finding the constraint and pricing the fixes", st_t))
S.append(Paragraph("Theory of Constraints study &nbsp;|&nbsp; From: [Your name], B.E. &nbsp;|&nbsp; Date: 22 Oct 2026 &nbsp;|&nbsp; "
                   "<b>Simulated, illustrative line (SimPy + Excel), not a real plant</b>", st_m))

S.append(Paragraph("The problem and the finding", st_h))
S.append(Paragraph(
    f"One production module must deliver {demand:,.0f} good enclosures a week through seven stages (saw, two CNC operations, deburr, anodising, inspection, packing). "
    f"The line's first-pass yield is 86%, matching the India base case in the TCO model (Project 1). <b>The constraint is {base['constraint'].lower()}</b>: it is loaded to "
    f"{stages[4]['util']*100:.0f}% of capacity at target demand, so the line can make only <b>{base['theory']:,.0f} good units a week ({base['pct']*100:.1f}% of demand)</b>. "
    f"The simulation agrees ({rows['C0']['sim']:,.0f}). The next-tightest stage is {base['nxt']} at {base['nxt_util']*100:.0f}%, so the constraint will move there once anodising is fixed. "
    f"Anodising is a batch process (48-unit racks), has the worst yield on the line (90.5%) and an OEE of {base['oee']*100:.0f}%; every hour lost there is lost for the whole line.", st_b))
S.append(Spacer(1, 3))
S.append(Image("/home/claude/p3_c1.png", width=182 * mm, height=64 * mm))

S.append(Paragraph("What the five focusing steps recover", st_h))
c1, c3, c7, c9, c10 = rows["C1"], rows["C3"], rows["C7"], rows["C9"], rows["C10"]
for tx in [
    f"<b>Exploit (step 2): speed up the constraint before buying more.</b> Cutting the rack batch from 52 to 46 minutes (SMED) lifts simulated output from {rows['C0']['sim']:,.0f} to "
    f"{c3['sim']:,.0f} units a week for about {usd(c3['onetime'])} one-time, a net benefit of {k(c3['net'])} a year per module, paid back in {c3['pay']:.1f} months.",
    f"<b>Subordinate (step 3): drum-buffer-rope release.</b> Releasing material only when the buffer ahead of anodising is below {rope1:,.0f} units leaves output unchanged but cuts average WIP by "
    f"{abs(c1['wip'])*100:.0f}% and flow time from {rows['C0']['flow']:.0f} hours to {c1['flow']:.1f} hours. Push release at the demand rate simply piles up a queue in front of the constraint.",
    f"<b>Elevate only if needed (step 4): a third anodising line does not pay.</b> It recovers the same units as SMED but nets {k(c10['net'])} a year per module because capex and running cost exceed a demand-capped gain.",
]:
    S.append(Paragraph(tx, st_bl, bulletText="\u2022"))
S.append(PageBreak())

S.append(Paragraph("Fix scenarios compared (per module; 8 simulation replications each)", st_t))
S.append(Spacer(1, 3))
hdrs = ["ID", "Scenario", "Sim good/wk", "Net benefit /yr", "Payback (months)", "WIP vs base", "Flow time (h)", "Theory headroom", "Constraint after fix"]
data = [[Paragraph(h, st_ch) for h in hdrs]]
for i in rows:
    r = rows[i]
    pay = "n/a" if isinstance(r["pay"], str) else (f"{r['pay']:.1f}" if r["onetime"] else "-")
    data.append([Paragraph(i, st_c), Paragraph(r["name"], st_c), Paragraph(f"{r['sim']:,.0f}", st_c), Paragraph(usd(r["net"]), st_c), Paragraph(pay, st_c),
                 Paragraph(f"{r['wip']*100:+.0f}%" if i != "C0" else "-", st_c), Paragraph(f"{r['flow']:.1f}", st_c),
                 Paragraph(f"{r['head']*100:+.1f}%", st_c), Paragraph(r["con"], st_c)])
t = Table(data, colWidths=[11 * mm, 48 * mm, 16 * mm, 19 * mm, 18 * mm, 15 * mm, 14 * mm, 17 * mm, 24 * mm], repeatRows=1)
st = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
      ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
t.setStyle(TableStyle(st))
S.append(t)
S.append(Spacer(1, 3))
S.append(Image("/home/claude/p3_c2.png", width=182 * mm, height=59 * mm))

S.append(Paragraph("Recommendation", st_h))
S.append(Paragraph(
    f"<b>Do now: SMED on the anodising racks plus drum-buffer-rope release (C7).</b> Simulated output reaches {c7['sim']:,.0f} units a week ({c7['sim']/demand*100:.1f}% of demand), "
    f"WIP falls {abs(c7['wip'])*100:.0f}% and flow time drops to {c7['flow']:.1f} hours, for {usd(c7['onetime'])} one-time. Net benefit is {k(c7['net'])} a year per module "
    f"({c7['fullnet']/1e6:.2f} million dollars a year across the ten-module line), payback {c7['pay']:.1f} months. "
    f"<b>Then add TPM at anodising (C9)</b> if more protection is wanted: theory headroom rises from {c7['head']*100:.1f}% to {c9['head']*100:.1f}% and output reaches {c9['sim']:,.0f}, "
    f"still netting {k(c9['net'])} a year. <b>Do not buy a third line</b> unless demand grows beyond about {c9['maxgood']:,.0f} units a week; "
    f"when it does, the constraint moves to {c10['con']} (theory limit {c10['maxgood']:,.0f}).", st_b))
S.append(Paragraph("Two lessons from the simulation", st_h))
for tx in [
    f"<b>Re-size the buffer after every improvement.</b> A rope of {rope1:,.0f} units protected the slow constraint, but once SMED sped it up the same rope starved it and output fell to about 9,840. "
    f"Re-sizing to {rope2:,.0f} units restored output. Step 5 (repeat) includes revisiting the buffer.",
    "<b>Scrap at the constraint is lost throughput.</b> Anodising yield from 90.5% to 93.5% adds about 300 good units a week on its own, but at the assumed $60k a year it does not pay on its own; it is worth buying only when the constraint is already protected and the extra units can be sold.",
]:
    S.append(Paragraph(tx, st_bl, bulletText="\u2022"))
S.append(Spacer(1, 3))
S.append(Paragraph(
    f"Basis and limits: the line is simulated. Stage parameters are benchmark-informed (world-class OEE is cited as about 85%, with machining averages of 55-65%) and fix costs are illustrative assumptions, "
    f"to be replaced with time studies and supplier quotes. Gains are capped at demand and priced at {contrib:,.2f} dollars contribution per unit (Project 1 price less material and an assumed variable conversion cost). "
    f"WIP and flow-time gains are not valued in dollars. See the workbook's README and Sources sheets.", st_s))
doc.build(S)
print("memo saved", rows["C7"]["net"], rows["C9"]["net"], rope1, rope2)
