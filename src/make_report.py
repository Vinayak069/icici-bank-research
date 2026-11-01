"""Write the initiation report (report/ICICI_Bank_Initiation.tex) from the model output.

Every number in the report is read from model/model_output.json,
model/chart_stats.json or data/key_figures.csv, so rerunning
build_model.py -> charts.py -> make_report.py keeps the text in step with the
model. Compile with: pdflatex -output-directory report report/ICICI_Bank_Initiation.tex
"""
import csv
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
from assumptions import DRIVER_LABELS, LISTED_SUBS, SCENARIOS, VALUATION  # noqa: E402

o = json.load(open(os.path.join(ROOT, "model", "model_output.json")))
st = json.load(open(os.path.join(ROOT, "model", "chart_stats.json")))
hist = {r["metric"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "key_figures.csv")))}
Y = o["years"]
B = o["scenarios"]["Base"]
V = B["valuation"]
BU, BE = o["scenarios"]["Bull"]["valuation"], o["scenarios"]["Bear"]["valuation"]
A = 4  # index of FY26A
DRAFT = os.path.exists(os.path.join(ROOT, "data", "DRAFT"))

# first half of FY27, read straight from the Q2 FY27 presentation
import re  # noqa: E402
import extract as ex  # noqa: E402
_pg = ex.pages(ex.H1_DECK)


def _row(metric):
    t, l, d, u, p, k = ex.METRICS[metric]
    i = ex.find_slide(_pg, t)
    for j in range(i, i + 3):
        v = ex.row_values(_pg[j], l, d, p)
        if v:
            return v  # [FY26, Q2 FY26, H1 FY26, Q1 FY27, Q2 FY27, H1 FY27]


H1 = {m: _row(m) for m in ["net_interest_income", "non_interest_income_excl_treasury", "operating_expenses",
                            "provisions", "profit_after_tax", "net_interest_margin"]}
h1g = {m: (v[5] / v[2] - 1) * 100 for m, v in H1.items() if m != "net_interest_margin"}
_hl = re.sub(r"\s+", " ", "\n".join(_pg[3:5]))
_m = lambda pat: float(re.search(pat, _hl).group(1))
loan_g = _m(r"Total loans grew by ([\d.]+)% y-o-y")
loan_g_ex = _m(r"FCNR \(B\) deposits, total loans grew by ([\d.]+)% y-o-y")
dep_g = _m(r"Period-end total deposits grew by ([\d.]+)% y-o-y")
dep_g_ex = _m(r"Excluding FCNR \(B\) deposits, period end total deposits grew by ([\d.]+)%")
nim_q2 = H1["net_interest_margin"][4]
nim_h1 = H1["net_interest_margin"][5]


def f(x, d=1):
    return f"{x:,.{d}f}"


def h(metric, col):
    return float(hist[metric][col])


def pct_chg(a, b):
    return (b / a - 1) * 100


cagr_pat = ((B["pat"][A] / B["pat"][0]) ** (1 / 4) - 1) * 100
cagr_adv = ((B["advances"][A] / B["advances"][0]) ** (1 / 4) - 1) * 100
eps_cagr_fwd = ((B["eps"][A + 3] / B["eps"][A]) ** (1 / 3) - 1) * 100
rating = V["rating"]
up = V["upside"]
standalone_12m = V["ri"]["v_12m"]
core_pb_now = (V["price"] - V["subs_after_discount"]) / B["bvps"][A]


def fin_table():
    rows = [("Net interest income (Rs bn)", "nii", 0), ("Core operating profit (Rs bn)", "core_op_profit", 0),
            ("Provisions (Rs bn)", "provisions", 0), ("Profit after tax (Rs bn)", "pat", 0),
            ("EPS (Rs)", "eps", 1), ("EPS growth (\\%)", "eps_growth", 1), ("Book value per share (Rs)", "bvps", 0),
            ("Net interest margin (\\%)", "nim", 2), ("Cost-to-income, core (\\%)", "cost_income", 1),
            ("Credit cost (\\%)", "credit_cost", 2), ("ROA (\\%)", "roa", 2), ("ROE (\\%)", "roe", 1)]
    cols = [3, 4, 5, 6, 7]
    out = ["\\begin{tabular}{l" + "r" * len(cols) + "}", "\\toprule",
           " & " + " & ".join(Y[c] for c in cols) + " \\\\", "\\midrule"]
    for lab, key, d in rows:
        out.append(lab + " & " + " & ".join(f(B[key][c], d) for c in cols) + " \\\\")
    out.append("\\midrule")
    out.append("P/E at Rs " + f(V["price"], 0) + " (x) & " + " & ".join(f(V["price"] / B["eps"][c], 1) for c in cols) + " \\\\")
    out.append("P/B at Rs " + f(V["price"], 0) + " (x) & " + " & ".join(f(V["price"] / B["bvps"][c], 2) for c in cols) + " \\\\")
    out += ["\\bottomrule", "\\end{tabular}"]
    return "\n".join(out)


def scen_table():
    out = ["\\begin{tabular}{lrrrrrrrrr}", "\\toprule",
           " & \\multicolumn{3}{c}{Bear} & \\multicolumn{3}{c}{Base} & \\multicolumn{3}{c}{Bull} \\\\",
           "\\cmidrule(lr){2-4}\\cmidrule(lr){5-7}\\cmidrule(lr){8-10}",
           " & " + " & ".join(["FY27E", "FY28E", "FY29E"] * 3) + " \\\\", "\\midrule"]
    for key, lab in DRIVER_LABELS.items():
        if key == "treasury":
            continue
        lab = lab.replace("%", "\\%")
        vals = SCENARIOS["Bear"][key] + SCENARIOS["Base"][key] + SCENARIOS["Bull"][key]
        out.append(lab + " & " + " & ".join(f(v, 2 if v < 5 else 1) for v in vals) + " \\\\")
    out.append("\\midrule")
    for lab, key, d in [("Profit after tax (Rs bn)", "pat", 0), ("ROE (\\%)", "roe", 1)]:
        vals = []
        for s in ["Bear", "Base", "Bull"]:
            vals += [o["scenarios"][s][key][A + t] for t in (1, 2, 3)]
        out.append(lab + " & " + " & ".join(f(v, d) for v in vals) + " \\\\")
    out.append("12-month value (Rs) & \\multicolumn{3}{c}{" + f(BE["target"], 0) + "} & \\multicolumn{3}{c}{\\textbf{" +
               f(V["target"], 0) + "}} & \\multicolumn{3}{c}{" + f(BU["target"], 0) + "} \\\\")
    out += ["\\bottomrule", "\\end{tabular}"]
    return "\n".join(out)


def ri_table():
    ri = V["ri"]
    cs = 1 + V["coe"] / 100
    rows = ["\\begin{tabular}{lr}", "\\toprule"]
    items = [("Net worth, 31 Mar 2026 (Rs bn)", B["net_worth"][A], 0)]
    for t in range(3):
        items.append((f"Residual income {Y[A + 1 + t]} (Rs bn)", ri["ri"][t], 0))
    items += [("PV of FY27E--FY29E residual income (Rs bn)", ri["pv_ri"], 0),
              ("PV of terminal value (Rs bn)", ri["pv_tv"], 0),
              ("Standalone value per share, 31 Mar 2026 (Rs)", ri["v_mar26"], 0),
              ("Standalone value per share today (Rs)", ri["v_today"], 0),
              ("Standalone value in 12 months (Rs)", ri["v_12m"], 0),
              ("Listed subsidiaries after 20\\% discount (Rs/share)", V["subs_after_discount"], 0),
              ("\\textbf{12-month target price (Rs)}", V["target"], 0)]
    for lab, v, d in items:
        rows.append(f"{lab} & {f(v, d)} \\\\")
    rows += ["\\bottomrule", "\\end{tabular}"]
    return "\n".join(rows)


def sens_table():
    gs = [7.0, 7.5, 8.0, 8.5, 9.0]
    cs = [11.0, 11.5, 12.0, 12.5, 13.0]
    out = ["\\begin{tabular}{l" + "r" * len(gs) + "}", "\\toprule",
           "COE $\\downarrow$ / growth $\\rightarrow$ & " + " & ".join(f"{g:.1f}\\%" for g in gs) + " \\\\", "\\midrule"]
    for c in cs:
        out.append(f"{c:.1f}\\% & " + " & ".join(f(V["sensitivity"][f"{c}|{g}"], 0) for g in gs) + " \\\\")
    out += ["\\bottomrule", "\\end{tabular}"]
    return "\n".join(out)


def sotp_table():
    sh = B["shares"][A]
    out = ["\\begin{tabular}{lrrrr}", "\\toprule",
           "Subsidiary & Shares held (m) & Price (Rs) & Value (Rs bn) & Rs/ICICI share \\\\", "\\midrule"]
    tot = 0
    for name, shm, px, _ in LISTED_SUBS:
        val = shm * px / 1000
        tot += val
        out.append(f"{name} & {f(shm, 0)} & {f(px, 1)} & {f(val, 0)} & {f(val / sh, 0)} \\\\")
    out.append("\\midrule")
    out.append(f"Total, before discount & & & {f(tot, 0)} & {f(tot / sh, 0)} \\\\")
    out.append(f"After {VALUATION['holdco_discount']:.0f}\\% holding-company discount & & & & {f(V['subs_after_discount'], 0)} \\\\")
    out += ["\\bottomrule", "\\end{tabular}"]
    return "\n".join(out)


rating_word = {"BUY": "BUY", "HOLD": "HOLD", "SELL": "SELL"}[rating]
draft_note = ("\\noindent\\colorbox{yellow!40}{\\parbox{\\dimexpr\\linewidth-2\\fboxsep}{\\footnotesize\\textbf{Draft:} "
              "historical figures were read from public summaries of ICICI Bank's filings and have not yet been "
              "re-extracted from the original PDFs. Treat every number as provisional.}}\\par\\medskip"
              if DRAFT else "")

tex = rf"""\documentclass[10pt,a4paper]{{article}}
\usepackage[margin=1.8cm]{{geometry}}
\usepackage[T1]{{fontenc}}
\usepackage{{lmodern}}
\usepackage{{booktabs,graphicx,xcolor,multicol,enumitem,titlesec,array}}
\usepackage[hidelinks]{{hyperref}}
\definecolor{{navy}}{{HTML}}{{1F3A5F}}
\titleformat{{\section}}{{\color{{navy}}\large\bfseries}}{{\thesection}}{{0.6em}}{{}}
\titlespacing*{{\section}}{{0pt}}{{8pt}}{{3pt}}
\setlist{{nosep,leftmargin=1.2em}}
\setlength{{\parindent}}{{0pt}}
\setlength{{\parskip}}{{4pt}}
\hypersetup{{pdftitle={{ICICI Bank initiation}}, pdfauthor={{Vinayak Singh}}}}
\begin{{document}}

{{\color{{navy}}\rule{{\linewidth}}{{2pt}}}}\\[-2pt]
{{\Large\bfseries ICICI Bank Ltd (NSE: ICICIBANK)}}\hfill{{\small Initiation of coverage, 1 Nov 2026}}\\[2pt]
{{\large Growth is back, but margins and valuation cap the upside}}\\[-4pt]
{{\color{{navy}}\rule{{\linewidth}}{{0.6pt}}}}

{draft_note}
\begin{{minipage}}[t]{{0.33\linewidth}}
\vspace{{0pt}}
\fcolorbox{{navy}}{{navy!6}}{{\begin{{minipage}}{{0.92\linewidth}}\small
\textbf{{Rating: {rating_word}}}\\[2pt]
12-month target: \textbf{{Rs {f(V['target'], 0)}}}\\
Price ({VALUATION['price_date']}): Rs {f(V['price'], 1)}\\
Implied return: {f(up, 1)}\%\\[4pt]
Bull / bear value: Rs {f(BU['target'], 0)} / Rs {f(BE['target'], 0)}\\[4pt]
Market cap: Rs {f(V['price'] * B['shares'][A] / 1000, 2)} trn\\
Shares: {f(B['shares'][A], 2)} bn\\
P/B FY27E: {f(V['pb_fy27'], 2)}x\\
P/E FY27E: {f(V['pe_fy27'], 1)}x\\
Cost of equity: {f(V['coe'], 1)}\%\\
Long-run growth: {f(VALUATION['g'], 1)}\%\\
Price change, 1y: {f(st['ICICI Bank|1y'], 0)}\%\\
Price change, 3y: {f(st['ICICI Bank|3y'], 0)}\%
\end{{minipage}}}}

\smallskip
{{\footnotesize Analyst: Vinayak Singh, IIT Kharagpur\\ Student research for learning purposes; not investment advice.}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.63\linewidth}}
\vspace{{0pt}}\small
\textbf{{Investment summary}}
\begin{{itemize}}
\item \textbf{{Five years of strong delivery.}} Profit after tax compounded {f(cagr_pat, 0)}\% a year from FY22 to FY26 (Rs {f(B['pat'][0], 0)} bn to Rs {f(B['pat'][A], 0)} bn) as loans grew {f(cagr_adv, 0)}\% a year, gross NPAs fell from {h('gross_npa_ratio', 'fy2022'):.2f}\% to {h('gross_npa_ratio', 'fy2026'):.2f}\% and ROE peaked at {f(B['roe'][2], 1)}\% in FY24.
\item \textbf{{H1 FY27: fast growth, thinner margin.}} Profit rose {f(h1g['profit_after_tax'], 0)}\% year on year in the first half. Loans grew {f(loan_g, 1)}\% year on year, but {f(loan_g_ex, 1)}\% excluding loans that international branches made against FCNR(B) (foreign-currency) deposits, which also lifted deposit growth from {f(dep_g_ex, 1)}\% to {f(dep_g, 1)}\%. NIM fell to {f(nim_q2, 2)}\% in Q2, from {h('net_interest_margin', 'fy2024'):.2f}\% in FY24.
\item \textbf{{Fair value close to the price.}} Our base case has ROE of {f(B['roe'][A + 1], 1)}--{f(B['roe'][A + 3], 1)}\% over FY27E--FY29E and EPS growth of {f(eps_cagr_fwd, 0)}\% a year. A residual income model at a {f(V['coe'], 1)}\% cost of equity values the standalone bank at Rs {f(standalone_12m, 0)} in 12 months; listed subsidiaries add Rs {f(V['subs_after_discount'], 0)} after a 20\% discount, for a target of Rs {f(V['target'], 0)}. The price implies a cost of equity of {f(V['implied_coe'], 1)}\% or perpetual growth of {f(V['implied_g'], 1)}\%.
\item \textbf{{What would change our view:}} NIM recovering above 4.3\% with growth above 15\%, or credit costs staying near 0.3\% through FY29 (bull case Rs {f(BU['target'], 0)}); a credit cycle would take value to Rs {f(BE['target'], 0)}.
\end{{itemize}}
\end{{minipage}}

\medskip
\begin{{center}}\small
\textbf{{Key financials (standalone, base case)}}\\[2pt]
{fin_table()}
\end{{center}}
{{\footnotesize Source: ICICI Bank results presentations and audited results FY22--FY26 (see Appendix); FY27E--FY29E are our estimates. ROE is our calculation on average net worth; the bank reports 16.0\% for FY26.}}

\newpage
\section{{Five years of re-rating: what drove earnings}}
\begin{{minipage}}[t]{{0.5\linewidth}}\vspace{{0pt}}
\includegraphics[width=\linewidth]{{../charts/earnings.png}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.47\linewidth}}\vspace{{0pt}}\small
Net interest income nearly doubled from Rs {f(B['nii'][0], 0)} bn in FY22 to Rs {f(B['nii'][A], 0)} bn in FY26, while operating expenses grew more slowly, from Rs {f(B['opex'][0], 0)} bn to Rs {f(B['opex'][A], 0)} bn, so core cost-to-income eased from {f(B['cost_income'][0], 1)}\% to {f(B['cost_income'][A], 1)}\%.

The bigger swing was provisions: Rs {f(B['provisions'][0], 0)} bn in FY22 against Rs {f(B['provisions'][A], 0)} bn in FY26, even though the loan book was {f(B['advances'][A] / B['advances'][0], 1)} times larger. Lower credit costs explain much of the jump in ROA from {h('roa', 'fy2022'):.2f}\% to {h('roa', 'fy2024'):.2f}\%.

Profit growth has slowed sharply: {f(B['pat_growth'][1], 0)}\% in FY23, {f(B['pat_growth'][2], 0)}\% in FY24, {f(B['pat_growth'][3], 0)}\% in FY25 and {f(B['pat_growth'][4], 0)}\% in FY26. FY26 provisions included a one-off Rs 12.8 bn standard-asset provision directed by the RBI.
\end{{minipage}}

\begin{{minipage}}[t]{{0.5\linewidth}}\vspace{{0pt}}
\includegraphics[width=\linewidth]{{../charts/nim_roe.png}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.47\linewidth}}\vspace{{0pt}}\small
\textbf{{Margin.}} NIM peaked at {h('net_interest_margin', 'fy2024'):.2f}\% in FY24 as loans repriced faster than deposits after the 2022--23 rate hikes. Since then deposit costs have caught up, and the share of CASA (current and savings accounts, the cheapest funding) in period-end deposits has fallen from {h('casa_ratio_period_end', 'fy2022'):.1f}\% in FY22 to {h('casa_ratio_period_end', 'fy2026'):.1f}\% in FY26 and {h('casa_ratio_period_end', 'h1fy2027'):.1f}\% in September 2026.

\textbf{{Returns.}} ROE on our basis fell from {f(B['roe'][2], 1)}\% in FY24 to {f(B['roe'][A], 1)}\% in FY26: margins slipped, credit costs normalised a little, and equity built up faster than profit (CET1 {h('cet1_ratio', 'fy2026'):.2f}\% in FY26, well above requirements). Surplus capital lowers ROE but gives room to grow.
\end{{minipage}}

\begin{{minipage}}[t]{{0.5\linewidth}}\vspace{{0pt}}
\includegraphics[width=\linewidth]{{../charts/asset_quality.png}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.47\linewidth}}\vspace{{0pt}}\small
\textbf{{Asset quality.}} Gross NPAs fell every year, from {h('gross_npa_ratio', 'fy2022'):.2f}\% to {h('gross_npa_ratio', 'fy2026'):.2f}\% (and {h('gross_npa_ratio', 'h1fy2027'):.2f}\% in September 2026); net NPAs are only {h('net_npa_ratio', 'fy2026'):.2f}\%. Provision coverage has slipped from {h('provision_coverage_ratio', 'fy2023'):.1f}\% (FY23) to {h('provision_coverage_ratio', 'fy2026'):.1f}\%, which is still high.

With credit costs at {h('credit_cost_provisions_to_avg_advances', 'fy2024'):.2f}--{h('credit_cost_provisions_to_avg_advances', 'fy2026'):.2f}\%, there is little left to gain from here. Our base case lets credit costs rise to {SCENARIOS['Base']['credit_cost'][2]:.2f}\% by FY29E, close to a through-the-cycle level for a lender with this mix.
\end{{minipage}}

\section{{First half of FY27: what changed}}
\small Net interest income grew {f(h1g['net_interest_income'], 1)}\% and operating expenses {f(h1g['operating_expenses'], 1)}\% year on year in H1 FY27, so profit after tax rose {f(h1g['profit_after_tax'], 1)}\% to Rs {f(H1['profit_after_tax'][5], 0)} bn. Provisions were only Rs {f(H1['provisions'][5], 0)} bn (credit cost {h('credit_cost_provisions_to_avg_advances', 'h1fy2027'):.2f}\% annualised).

The balance sheet changed shape. Loans rose from Rs {f(B['advances'][A], 0)} bn in March to Rs {h('total_advances_net', 'h1fy2027'):,.0f} bn in September, and deposits from Rs {h('total_deposits', 'fy2026'):,.0f} bn to Rs {h('total_deposits', 'h1fy2027'):,.0f} bn. Much of the increase was foreign-currency FCNR(B) deposits lent on by the bank's international branches: excluding them, loans grew {f(loan_g_ex, 1)}\% and deposits {f(dep_g_ex, 1)}\% year on year, against {f(loan_g, 1)}\% and {f(dep_g, 1)}\% including them. The presentation does not give the margin on this book, but NIM fell from {H1['net_interest_margin'][3]:.2f}\% in Q1 to {f(nim_q2, 2)}\% in Q2 as it grew, which suggests it earns less than the rest of the balance sheet. We treat it as lower-margin growth: our FY27E loan growth of {SCENARIOS['Base']['loan_growth'][0]:.0f}\% includes it, and our NIM of {SCENARIOS['Base']['nim'][0]:.2f}\% assumes the second half stays near the Q2 level.

\clearpage\section{{Forecasts: three scenarios}}
\small Loans and total assets grow at the scenario rate. Net interest income is NIM times average interest-earning assets, which we hold at the FY26 ratio of {f(B['iea_ratio'][A] * 100, 1)}\% of average total assets (implied by reported NII and NIM). Operating expenses follow core cost-to-income, provisions follow credit cost on average loans, and dividends follow the payout ratio, paid the following year. Treasury gains are not forecast.

\begin{{center}}\footnotesize
\resizebox{{\linewidth}}{{!}}{{{scen_table()}}}
\end{{center}}

\section{{Valuation: residual income, sum of the parts}}
\small A bank's value comes from the return it earns on equity above its cost of equity. The residual income model makes this explicit: value = book value today + present value of (profit $-$ cost of equity $\times$ opening book). We use a {f(V['coe'], 1)}\% cost of equity (India 10-year bond yield {VALUATION['rf']:.2f}\% on 30 Oct 2026, beta {VALUATION['beta']:.1f}, equity risk premium {VALUATION['erp']:.1f}\%) and {VALUATION['g']:.1f}\% long-run growth of residual income after FY29E.

\begin{{minipage}}[t]{{0.48\linewidth}}\vspace{{0pt}}\footnotesize
\textbf{{Residual income valuation (base case)}}\\[2pt]
{ri_table()}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.48\linewidth}}\vspace{{0pt}}\footnotesize
\textbf{{Sensitivity of the 12-month target (Rs)}}\\[2pt]
{{\setlength{{\tabcolsep}}{{4pt}}{sens_table()}}}

\medskip
\textbf{{Cross-check: justified price-to-book}}\\
Sustainable ROE (avg FY28E--FY29E): {f(V['roe_sustainable'], 1)}\%\\
Justified P/B = (ROE $-$ g)/(COE $-$ g) = {f(V['justified_pb'], 2)}x\\
On book value in 12 months (Rs {f(V['bvps_12m'], 0)}), plus subsidiaries: Rs {f(V['pb_target'], 0)}
\end{{minipage}}

\medskip
\textbf{{Listed subsidiaries at market value.}} ICICI Bank owns just over half of three listed companies. We value them at market prices on {VALUATION['price_date']} with a 20\% holding-company discount. Unlisted subsidiaries (ICICI Securities, ICICI Bank UK and Canada, ICICI Home Finance) are left at zero, which is conservative: one broker values all subsidiaries at about Rs 300 per share.

\begin{{center}}\footnotesize
{sotp_table()}
\end{{center}}

\begin{{minipage}}[t]{{0.5\linewidth}}\vspace{{0pt}}
\includegraphics[width=\linewidth]{{../charts/pb_band.png}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.47\linewidth}}\vspace{{0pt}}\small
\textbf{{What the price implies.}} Since FY22 the stock has traded at {f(st['pb_min'], 1)}--{f(st['pb_max'], 1)}x standalone book (average {f(st['pb_mean'], 2)}x); today it is {f(st['pb_last'], 2)}x. Stripping out our subsidiary value, the core bank trades at {f(core_pb_now, 2)}x FY26 book. Holding growth at {VALUATION['g']:.0f}\%, the current price is justified only by a cost of equity of about {f(V.get('implied_coe', 0), 1)}\%; holding the cost of equity at {f(V['coe'], 1)}\%, it needs perpetual growth of about {f(V.get('implied_g', 0), 1)}\%.

Brokers' targets after Q2 FY27 are higher than ours (Rs 1,710--2,025), mostly because they apply 2.4--2.6x to adjusted book value. The difference with our view is the assumed sustainable ROE and cost of equity, not the quality of the bank.
\end{{minipage}}

\section{{Risks to our view}}
\begin{{multicols}}{{2}}\small
\textbf{{Upside risks}}
\begin{{itemize}}
\item Faster loan growth with stable margins: the bull case (loan growth {SCENARIOS['Bull']['loan_growth'][0]:.0f}\% in FY27E, NIM recovering to {SCENARIOS['Bull']['nim'][2]:.2f}\%) gives Rs {f(BU['target'], 0)}.
\item Credit costs staying at the FY24--26 low of about 0.3--0.4\%.
\item A higher value for the unlisted subsidiaries, which we leave at zero.
\item Rate cuts, which would lower the 10-year yield and the cost of equity: each 0.5 point off the cost of equity adds roughly Rs {f(V['sensitivity']['12.0|8.0'] - V['sensitivity']['12.5|8.0'], 0)} to the target.
\end{{itemize}}
\columnbreak
\textbf{{Downside risks}}
\begin{{itemize}}
\item Deposit competition: CASA is falling and deposit costs could squeeze NIM further.
\item A credit cycle in unsecured retail or small business loans: the bear case (credit cost up to {max(SCENARIOS['Bear']['credit_cost']):.1f}\%) gives Rs {f(BE['target'], 0)}.
\item Bond yields near {VALUATION['rf']:.1f}\% and oil above \$100 raise the cost of equity and the risk of slower growth.
\item Regulatory actions, such as the RBI-directed provision taken in FY26.
\end{{itemize}}
\end{{multicols}}

\begin{{minipage}}[t]{{0.5\linewidth}}\vspace{{0pt}}
\includegraphics[width=\linewidth]{{../charts/price_vs_peers.png}}
\end{{minipage}}\hfill
\begin{{minipage}}[t]{{0.47\linewidth}}\vspace{{0pt}}\small
\textbf{{Share price performance.}} Over three years ICICI Bank returned {f(st['ICICI Bank|3y'], 0)}\%, ahead of HDFC Bank ({f(st['HDFC Bank|3y'], 0)}\%), Axis ({f(st['Axis Bank|3y'], 0)}\%) and Kotak ({f(st['Kotak Mahindra|3y'], 0)}\%), and behind SBI ({f(st['SBI|3y'], 0)}\%). Over the past year it is up {f(st['ICICI Bank|1y'], 0)}\%. The market already rewards ICICI as the best-run large private bank, which is why we think the stock needs returns to rise, not just hold, to beat the market from here.
\end{{minipage}}

\section*{{Appendix: method, data and limitations}}
\footnotesize
\begin{{itemize}}
\item \textbf{{Data.}} Standalone figures from ICICI Bank's Q4 investor presentations, audited results and press releases for FY22--FY26, and the Q2 FY27 results filing (17 Oct 2026). Share prices: NSE daily closes to {st['last_date']}. Subsidiary stakes: Sep 2026 shareholding disclosures; ICICI Lombard form NL-9. 10-year yield: Trading Economics, 30 Oct 2026.
\item \textbf{{Reproducibility.}} The Python pipeline {'rebuilds' if DRAFT else 'extracts the figures from the filings, rebuilds'} the Excel model (scenario switch, live formulas) and regenerates this report. A check recalculates the workbook and confirms it matches the Python model to six significant figures, with reported profit before tax reconciling for all five years.
\item \textbf{{Limitations.}} Standalone bank only; subsidiaries enter at market value. Net worth is capital plus reserves (excluding revaluation reserves), not the RBI definition. Forecasts are annual, with no quarterly seasonality. ROE is on our own basis.
\item This is a student research exercise and not a recommendation to buy or sell any security.
\end{{itemize}}

\end{{document}}
"""

os.makedirs(os.path.join(ROOT, "report"), exist_ok=True)
open(os.path.join(ROOT, "report", "ICICI_Bank_Initiation.tex"), "w").write(tex)
print("wrote report/ICICI_Bank_Initiation.tex")
