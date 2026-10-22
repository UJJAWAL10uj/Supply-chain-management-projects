"""One-page sourcing memo; every number is read from the recalculated Excel model."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

XLSX = "/mnt/user-data/outputs/Apple_Sourcing_TCO_Model_Aluminium_Enclosure.xlsx"
OUT = "/mnt/user-data/outputs/Sourcing_Recommendation_Memo_Aluminium_Enclosure.pdf"

wb = load_workbook(XLSX, data_only=True)
S, T, I = wb["Summary"], wb["Sensitivity"], wb["Incoterms"]

# scenario table
sc = {}
for r in range(6, 17):
    sc[S[f"A{r}"].value] = dict(india=S[f"B{r}"].value, china=S[f"C{r}"].value, viet=S[f"D{r}"].value,
                                cheapest=S[f"E{r}"].value, gap=S[f"F{r}"].value)
base = sc["Base case"]
names = list(sc.keys())
shock = sc["Import-shock stack (duty +10, freight x2.5, Al +15%)"]
irisk = sc["India-risk stack (vol -10%, ECMS 0%, yield -4 pts)"]

be_bcd, be_frt, be_yield = S["B23"].value, S["B24"].value, S["B26"].value
base_bcd = wb["Inputs"]["D36"].value

# cost stack (rows 30-42, cols B,C,D = India, China, Vietnam)
stack = {S[f"A{r}"].value: (S[f"B{r}"].value, S[f"C{r}"].value, S[f"D{r}"].value) for r in range(30, 43)}
yl_i, yl_c = stack["Yield loss"][0], stack["Yield loss"][1]
mc_i, mc_c = stack["Machining"][0], stack["Machining"][1]
penalty = (yl_i - yl_c) + (mc_i - mc_c)
imp_burden = sum(stack[k][1] for k in ["International freight", "Cargo insurance", "Import duties (BCD + SWS + other)",
                                       "Clearance and haulage", "Pipeline carrying cost"])
ecms = -stack["ECMS incentive to buyer"][0]

# spreads
def spread(key):
    vals = [v[key] for v in sc.values()]
    return min(vals), max(vals), max(vals) / min(vals) - 1
ind_lo, ind_hi, ind_sp = spread("india")
chn_lo, chn_hi, chn_sp = spread("china")

# yield / ECMS grid: rows 40-45 yields 0.80..0.95 ; cols B..F pass-through 0..1.0
adv_86_100 = T["F42"].value      # yield 86%, pass-through 100%
adv_92_100 = T["F44"].value      # yield 92%, pass-through 100%

# incoterms
fob, cif, ddp = I["C16"].value, I["E16"].value, I["G16"].value
ddp_prem, cif_prem = I["G17"].value, I["E17"].value

# ---- chart (gap % by scenario)
short = {
    "Base case": "Base case", "Import duty +10 pts": "Duty +10 pts", "Import duty exempt (BCD to 0)": "Duty exempt",
    "Freight spike x2.5": "Freight x2.5", "Volume -10%": "Volume -10%", "Volume +10%": "Volume +10%",
    "Aluminium +15%": "Aluminium +15%", "ECMS delayed (0% pass-through)": "ECMS delayed",
    "India yield -4 pts": "India yield -4 pts",
    "Import-shock stack (duty +10, freight x2.5, Al +15%)": "Import-shock stack",
    "India-risk stack (vol -10%, ECMS 0%, yield -4 pts)": "India-risk stack",
}
labels = [short[n] for n in names]
gaps = [sc[n]["gap"] * 100 for n in names]
fig, ax = plt.subplots(figsize=(3.6, 2.9), dpi=200)
cols = ["#C0504D" if g > 0 else "#4F9A5B" for g in gaps]
ax.barh(range(len(labels)), gaps, color=cols, height=0.62)
ax.set_yticks(range(len(labels)))
ax.set_yticklabels(labels, fontsize=6)
ax.invert_yaxis()
ax.axvline(0, color="black", lw=0.6)
ax.tick_params(axis="x", labelsize=6)
ax.set_xlabel("India local vs China import, landed cost (%)", fontsize=6)
ax.set_title("Red: India dearer   Green: India cheaper", fontsize=6.5)
for sp in ("top", "right"):
    ax.spines[sp].set_visible(False)
plt.tight_layout()
plt.savefig("/home/claude/gap_chart.png")
plt.close()

# ---- PDF
NAVY = colors.HexColor("#1F3864")
st_title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=NAVY)
st_meta = ParagraphStyle("m", fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#444444"))
st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=NAVY, spaceBefore=4, spaceAfter=1)
st_b = ParagraphStyle("b", fontName="Helvetica", fontSize=8.2, leading=10.4)
st_bl = ParagraphStyle("bl", parent=st_b, leftIndent=9, bulletIndent=0)
st_small = ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=7, leading=8.6, textColor=colors.HexColor("#555555"))

usd = lambda v: f"${v:,.2f}"
pct = lambda v: f"{v*100:+.1f}%"

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=11 * mm, bottomMargin=9 * mm,
                        title="Sourcing recommendation: CNC aluminium enclosure", author="Independent project")
story = []
story.append(Paragraph("Sourcing recommendation: CNC aluminium enclosure for final assembly near Chennai", st_title))
story.append(Paragraph("To: Global Sourcing &amp; Supply team &nbsp;|&nbsp; From: [Your name], B.E. &nbsp;|&nbsp; Date: 22 Oct 2026 &nbsp;|&nbsp; "
                       "Basis: attached Excel TCO model (independent project, public data plus stated assumptions)", st_meta))
story.append(Spacer(1, 3))

story.append(Paragraph("Recommendation", st_h))
story.append(Paragraph(
    f"<b>Qualify an Indian supplier as a second source now, but do not move the bulk of volume on cost alone yet.</b> "
    f"At current duty, freight and yield assumptions, a China import lands at <b>{usd(base['china'])}</b> per unit versus "
    f"<b>{usd(base['india'])}</b> from a local Indian supplier (<b>{pct(base['gap'])}</b>); Vietnam lands at {usd(base['viet'])}. "
    f"The local option becomes cheapest only when imports are hit: import duty at or above about <b>{be_bcd*100:.0f}%</b> "
    f"(modelled today: {base_bcd*100:.0f}%), freight above about <b>{be_frt:.1f}x</b>, or the combined import-shock case "
    f"(India {abs(shock['india']/shock['china']-1)*100:.1f}% cheaper).", st_b))

story.append(Paragraph("What drives the gap", st_h))
for t in [
    f"<b>Yield and machining, not logistics.</b> India carries a {usd(penalty)}/unit penalty on yield loss ({usd(yl_i)} vs {usd(yl_c)}) and machining "
    f"({usd(mc_i)} vs {usd(mc_c)}). China's entire import burden (freight, duty, handling, carrying cost) is {usd(imp_burden)}/unit.",
    f"<b>The ECMS incentive helps but does not close it alone.</b> At 50% pass-through it is worth {usd(ecms)}/unit. At 86% yield, even full pass-through "
    f"leaves India {usd(abs(adv_86_100))}/unit dearer; parity needs about {be_yield*100:.0f}% yield at 50% pass-through, or about 92% with full pass-through "
    f"({'India ahead by ' + usd(adv_92_100) if adv_92_100 > 0 else 'still short by ' + usd(abs(adv_92_100))}).",
    f"<b>Local supply is the stable hedge.</b> India's landed cost spans {usd(ind_lo)} to {usd(ind_hi)} across all 11 scenarios ({ind_sp*100:.0f}% spread) "
    f"versus {usd(chn_lo)} to {usd(chn_hi)} for China ({chn_sp*100:.0f}% spread). Volume swings of 10% move cost by only about 1%.",
]:
    story.append(Paragraph(t, st_bl, bulletText="\u2022"))

# scenario table + chart side by side
pick = ["Base case", "Import duty +10 pts", "Freight spike x2.5", "India yield -4 pts",
        "Import-shock stack (duty +10, freight x2.5, Al +15%)", "India-risk stack (vol -10%, ECMS 0%, yield -4 pts)"]
rows = [["Scenario (USD per unit)", "India", "China", "Vietnam", "Cheapest"]]
for n in pick:
    d = sc[n]
    rows.append([short[n], usd(d["india"]), usd(d["china"]), usd(d["viet"]), d["cheapest"]])
tbl = Table(rows, colWidths=[33 * mm, 15 * mm, 15 * mm, 16 * mm, 17 * mm])
tbl.setStyle(TableStyle([
    ("FONT", (0, 0), (-1, -1), "Helvetica", 7.3),
    ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7.3),
    ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
    ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
]))
img = Image("/home/claude/gap_chart.png", width=70 * mm, height=56 * mm)
side = Table([[tbl, img]], colWidths=[98 * mm, 80 * mm])
side.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
story.append(Spacer(1, 3))
story.append(side)

story.append(Paragraph("Incoterms choice for imported supply", st_h))
story.append(Paragraph(
    f"For a China import the buyer-managed FOB cost is {usd(fob)}; CIF adds {cif_prem*100:.1f}% and DDP adds {ddp_prem*100:.1f}% "
    f"({usd(ddp - fob)}/unit). DDP is inexpensive protection if the team wants the seller to carry clearance work and duty-change handling, "
    f"but suppliers typically re-price when the duty rate moves, so duty exposure is reduced rather than removed.", st_b))

story.append(Paragraph("Recommended next steps", st_h))
for t in [
    "<b>Replace assumptions with quotes:</b> RFQ two or three Indian CNC suppliers for cycle time, first-pass yield and tooling, and re-run the model.",
    "<b>Stage volume against yield milestones:</b> release additional share only as measured first-pass yield approaches the break-even range above.",
    "<b>Lock the incentive terms:</b> negotiate the share of ECMS benefit passed to the buyer in the supply agreement.",
    "<b>Verify tariff inputs:</b> confirm the current Indian duty on the enclosure's HS code and any preferential rate for Vietnam-origin parts.",
]:
    story.append(Paragraph(t, st_bl, bulletText="\u2022"))

story.append(Spacer(1, 3))
story.append(Paragraph(
    "Context not modelled: US duties on finished devices differ by assembly country (Section 301 forced-labor duty from 24 Jul 2026: India 10%, China and Vietnam 12.5%, "
    "plus China's older list duties), which adds a strategic reason to localise. Inputs: LME aluminium USD 3,134/t (6 Oct 2026) and the ECMS enclosure incentive range of 6-8% "
    "are sourced; cycle times, rates, yields, tooling and freight are illustrative assumptions, not Apple data or supplier quotes. See the Sources sheet.", st_small))

doc.build(story)
print("memo saved")
print(dict(base=base, be_bcd=be_bcd, be_frt=be_frt, be_yield=be_yield, penalty=penalty, imp_burden=imp_burden, ecms=ecms,
           ind=(ind_lo, ind_hi, ind_sp), chn=(chn_lo, chn_hi, chn_sp), adv86=adv_86_100, adv92=adv_92_100,
           fob=fob, cif_prem=cif_prem, ddp_prem=ddp_prem, shock_gap=shock['india']/shock['china']-1))
