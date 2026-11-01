"""Build the ICICI Bank forecast and valuation model.

Reads the five-year history (data/key_figures.csv), forecasts FY27E-FY29E
from src/assumptions.py, values the bank with a residual income model and a
justified price-to-book cross-check, and writes:

  model/ICICI_Bank_Model.xlsx   live Excel model (formulas, scenario switch)
  model/model_output.json       the same numbers computed in Python, used by
                                the charts, the report and check_model.py

All money is Rs billion unless stated; per-share figures are Rs.
"""
import csv
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
from assumptions import DRIVER_LABELS, LISTED_SUBS, SCENARIOS, VALUATION, YEARS_E  # noqa: E402

YEARS_A = ["FY22A", "FY23A", "FY24A", "FY25A", "FY26A"]
CSV_COLS = ["fy2022", "fy2023", "fy2024", "fy2025", "fy2026"]
FACE_VALUE = 2.0  # Rs per share

HIST_METRICS = [
    ("net_interest_income", "Net interest income"),
    ("non_interest_income_excl_treasury", "Non-interest income ex treasury"),
    ("fee_income", "  of which fee income"),
    ("treasury_income", "Treasury income"),
    ("operating_expenses", "Operating expenses"),
    ("core_operating_profit", "Core operating profit (reported)"),
    ("provisions", "Provisions"),
    ("profit_before_tax", "Profit before tax (reported)"),
    ("tax", "Tax"),
    ("profit_after_tax", "Profit after tax"),
    ("total_advances_net", "Advances (net)"),
    ("total_deposits", "Deposits"),
    ("total_assets", "Total assets"),
    ("equity_share_capital", "Equity share capital"),
    ("net_worth", "Net worth (capital + reserves)"),
    ("net_interest_margin", "Net interest margin (%)"),
    ("gross_npa_ratio", "Gross NPA (%)"),
    ("net_npa_ratio", "Net NPA (%)"),
    ("provision_coverage_ratio", "Provision coverage (%)"),
    ("cet1_ratio", "CET1 ratio (%)"),
    ("total_car", "Capital adequacy (%)"),
    ("roa", "ROA reported (%)"),
    ("eps_basic", "EPS, weighted average, reported (Rs)"),
    ("book_value_per_share", "Book value per share, reported (Rs)"),
    ("dividend_per_share", "Dividend per share (Rs)"),
]


def load_history(path):
    rows = {}
    with open(path) as f:
        for r in csv.DictReader(f):
            rows[r["metric"]] = r
    hist, src = {}, {}
    for key, _ in HIST_METRICS:
        r = rows[key]
        vals = []
        for c in CSV_COLS:
            v = (r.get(c) or "").strip()
            vals.append(float(v) if v else None)
        hist[key] = vals
        src[key] = "; ".join(sorted({r.get(f"source_{c}", "").split(" p")[0] + " p" + r.get(f"source_{c}", "").split(" p")[-1]
                                     for c in CSV_COLS if r.get(f"source_{c}")}))
    return hist, src


# ---------------------------------------------------------------- python model
def run_model(hist, drv, val):
    """Return a dict of series, each 8 long (5 actual + 3 forecast)."""
    n_a = len(YEARS_A)
    m = {k: list(v) for k, v in hist.items()}

    adv = m["total_advances_net"][:]
    assets = m["total_assets"][:]
    nii = m["net_interest_income"][:]
    nonii = m["non_interest_income_excl_treasury"][:]
    trea = m["treasury_income"][:]
    opex = m["operating_expenses"][:]
    prov = m["provisions"][:]
    tax = m["tax"][:]
    nim = m["net_interest_margin"][:]
    shares = [c / FACE_VALUE for c in m["equity_share_capital"]]
    nw = m["net_worth"][:]
    dps = m["dividend_per_share"][:]

    # Interest-earning assets implied by reported NII / NIM, as a share of
    # average total assets. FY26 ratio carries forward.
    avg_assets = [None] + [(assets[i - 1] + assets[i]) / 2 for i in range(1, n_a)]
    iea_ratio = [None] + [nii[i] / (nim[i] / 100) / avg_assets[i] for i in range(1, n_a)]

    for t in range(3):
        i = n_a + t
        adv.append(adv[i - 1] * (1 + drv["loan_growth"][t] / 100))
        assets.append(assets[i - 1] * (1 + drv["loan_growth"][t] / 100))
        avg_assets.append((assets[i - 1] + assets[i]) / 2)
        iea_ratio.append(iea_ratio[n_a - 1])
        nim.append(drv["nim"][t])
        nii.append(nim[i] / 100 * iea_ratio[i] * avg_assets[i])
        nonii.append(nonii[i - 1] * (1 + drv["nonii_growth"][t] / 100))
        trea.append(drv["treasury"][t])
        opex.append(drv["cost_income"][t] / 100 * (nii[i] + nonii[i]))
        prov.append(drv["credit_cost"][t] / 100 * (adv[i - 1] + adv[i]) / 2)
        pbt_i = nii[i] + nonii[i] - opex[i] - prov[i] + trea[i]
        tax.append(drv["tax_rate"][t] / 100 * pbt_i)
        shares.append(shares[n_a - 1])
    pbt = [nii[i] + nonii[i] - opex[i] - prov[i] + trea[i] for i in range(n_a + 3)]
    pat = [pbt[i] - tax[i] for i in range(n_a + 3)]
    for t in range(3):
        i = n_a + t
        eps_i = pat[i] / shares[i]
        dps.append(drv["payout"][t] / 100 * eps_i)
        # dividends paid in year i are last year's DPS
        nw.append(nw[i - 1] + pat[i] - dps[i - 1] * shares[i - 1])

    core_inc = [nii[i] + nonii[i] for i in range(n_a + 3)]
    cop = [core_inc[i] - opex[i] for i in range(n_a + 3)]
    ci = [opex[i] / core_inc[i] * 100 for i in range(n_a + 3)]
    avg_adv = [None] + [(adv[i - 1] + adv[i]) / 2 for i in range(1, n_a + 3)]
    cc = [None] + [prov[i] / avg_adv[i] * 100 for i in range(1, n_a + 3)]
    avg_nw = [None] + [(nw[i - 1] + nw[i]) / 2 for i in range(1, n_a + 3)]
    roe = [None] + [pat[i] / avg_nw[i] * 100 for i in range(1, n_a + 3)]
    roa = [None] + [pat[i] / avg_assets[i] * 100 for i in range(1, n_a + 3)]
    eps = [pat[i] / shares[i] for i in range(n_a + 3)]
    bvps = [nw[i] / shares[i] for i in range(n_a + 3)]
    tax_rate = [tax[i] / pbt[i] * 100 for i in range(n_a + 3)]
    growth = lambda s: [None] + [(s[i] / s[i - 1] - 1) * 100 for i in range(1, len(s))]

    out = dict(advances=adv, adv_growth=growth(adv), assets=assets, avg_assets=avg_assets,
               iea_ratio=iea_ratio, nim=nim, nii=nii, nii_growth=growth(nii), nonii=nonii,
               treasury=trea, core_income=core_inc, opex=opex, cost_income=ci,
               core_op_profit=cop, provisions=prov, credit_cost=cc, pbt=pbt, tax=tax,
               tax_rate=tax_rate, pat=pat, pat_growth=growth(pat), net_worth=nw,
               roe=roe, roa=roa, shares=shares, eps=eps, eps_growth=growth(eps),
               bvps=bvps, dps=dps)
    out["valuation"] = value(out, val)
    return out


def ri_value_per_share(o, coe, g, val):
    """Residual income value per share at 31 Mar 2026, today, and 12 months out."""
    n_a = len(YEARS_A)
    c, gg = coe / 100, g / 100
    nw, pat, sh = o["net_worth"], o["pat"], o["shares"][n_a - 1]
    pv = 0.0
    ri = []
    for t in range(3):
        i = n_a + t
        r = pat[i] - c * nw[i - 1]
        ri.append(r)
        pv += r / (1 + c) ** (t + 1)
    tv = ri[-1] * (1 + gg) / (c - gg)
    pv_tv = tv / (1 + c) ** 3
    v_mar26 = (nw[n_a - 1] + pv + pv_tv) / sh
    v_today = v_mar26 * (1 + c) ** (val["months_since_fy26"] / 12) - val["dps_paid_since_fy26"]
    v_12m = v_today * (1 + c) - o["dps"][n_a]
    return dict(ri=ri, pv_ri=pv, tv=tv, pv_tv=pv_tv, v_mar26=v_mar26, v_today=v_today, v_12m=v_12m)


def value(o, val):
    n_a = len(YEARS_A)
    coe = val["rf"] + val["beta"] * val["erp"]
    ri = ri_value_per_share(o, coe, val["g"], val)
    subs_bn = sum(sh * px / 1000 for _, sh, px, _ in LISTED_SUBS)
    subs_gross = subs_bn / o["shares"][n_a - 1]
    subs = subs_gross * (1 - val["holdco_discount"] / 100)
    target = ri["v_12m"] + subs
    # justified P/B cross-check on book value 12 months out (Oct 2027)
    roe_sus = (o["roe"][n_a + 1] + o["roe"][n_a + 2]) / 2
    pb = (roe_sus - val["g"]) / (coe - val["g"])
    frac = (val["months_since_fy26"] + 12 - 12) / 12  # months into FY28
    bv_12m = o["bvps"][n_a] + frac * (o["bvps"][n_a + 1] - o["bvps"][n_a])
    pb_target = pb * bv_12m + subs
    price = val["price"]
    upside = (target / price - 1) * 100
    rating = "BUY" if upside > 15 else ("SELL" if upside < -5 else "HOLD")
    sens = {}
    for c in [11.0, 11.5, 12.0, 12.5, 13.0]:
        for g in [7.0, 7.5, 8.0, 8.5, 9.0]:
            sens[f"{c}|{g}"] = ri_value_per_share(o, c, g, val)["v_12m"] + subs
    # what the current price implies: cost of equity at our growth, and growth at our cost of equity
    def solve(fn, lo, hi, increasing):
        for _ in range(80):
            mid = (lo + hi) / 2
            above = fn(mid) > price
            if above == increasing:
                hi = mid
            else:
                lo = mid
        return (lo + hi) / 2
    price = val["price"]
    implied_coe = solve(lambda c: ri_value_per_share(o, c, val["g"], val)["v_today"] + subs,
                        val["g"] + 0.5, 20.0, increasing=False)
    implied_g = solve(lambda g: ri_value_per_share(o, coe, g, val)["v_today"] + subs,
                      0.0, coe - 0.2, increasing=True)
    return dict(coe=coe, ri=ri, implied_coe=implied_coe, implied_g=implied_g, subs_value_bn=subs_bn, subs_gross_per_share=subs_gross,
                subs_after_discount=subs, target=target,
                roe_sustainable=roe_sus, justified_pb=pb, bvps_12m=bv_12m,
                pb_target=pb_target, price=price, upside=upside, rating=rating,
                implied_pb_now=price / (o["bvps"][n_a - 1] + frac * 0),
                pe_fy27=price / o["eps"][n_a], pe_fy28=price / o["eps"][n_a + 1],
                pb_fy27=price / o["bvps"][n_a], pb_fy28=price / o["bvps"][n_a + 1],
                sensitivity=sens)


# ---------------------------------------------------------------- excel model
BLUE = Font(color="0000CC")
BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="DDE5F0")
INPUT = PatternFill("solid", fgColor="FFF7D6")
THIN = Border(bottom=Side(style="thin", color="999999"))
PCT = "0.00"
BN = "#,##0.0"


def write_excel(hist, src, path):
    wb = Workbook()
    n_a = len(YEARS_A)

    # ------------------------------------------------------------ Assumptions
    wa = wb.active
    wa.title = "Assumptions"
    wa["A1"] = "ICICI Bank: forecast drivers and valuation inputs"
    wa["A1"].font = Font(bold=True, size=13)
    wa["A2"] = "Scenario (1 = Base, 2 = Bull, 3 = Bear)"
    wa["C2"] = 1
    wa["C2"].fill = INPUT
    wa["C2"].font = BLUE
    wa["A3"] = "Yellow cells are inputs. Change C2 to switch every forecast and the valuation."
    hdr = 5
    for j, h in enumerate(["Driver", "Scenario"] + YEARS_E):
        cell = wa.cell(hdr, 1 + j, h)
        cell.font = BOLD
        cell.fill = HEAD
    active = {}
    r = hdr + 1
    for key, label in DRIVER_LABELS.items():
        first = r
        for sname in ["Base", "Bull", "Bear"]:
            wa.cell(r, 1, label if sname == "Base" else "")
            wa.cell(r, 2, sname)
            for t in range(3):
                c = wa.cell(r, 3 + t, SCENARIOS[sname][key][t])
                c.fill = INPUT
                c.font = BLUE
            r += 1
        wa.cell(r, 2, "Active").font = BOLD
        for t in range(3):
            col = get_column_letter(3 + t)
            wa.cell(r, 3 + t, f"=CHOOSE($C$2,{col}{first},{col}{first + 1},{col}{first + 2})").font = BOLD
        active[key] = r
        r += 2
    vr = r + 1
    wa.cell(vr, 1, "Valuation inputs").font = Font(bold=True, size=12)
    vrows = {}
    vitems = [("rf", "Risk-free rate, India 10Y G-sec (%)"), ("beta", "Beta"),
              ("erp", "Equity risk premium (%)"), ("coe", "Cost of equity (%)"),
              ("g", "Long-run growth (%)"), ("price", "Share price (Rs)"),
              ("price_date", "Price date"), ("dps_paid_since_fy26", "Dividend paid since Mar-26 (Rs)"),
              ("months_since_fy26", "Months from 31 Mar 2026 to valuation date"),
              ("holdco_discount", "Holding-company discount (%)")]
    rr = vr + 1
    for key, label in vitems:
        wa.cell(rr, 1, label)
        if key == "coe":
            wa.cell(rr, 3, f"=C{vrows['rf']}+C{vrows['beta']}*C{vrows['erp']}").font = BOLD
        else:
            c = wa.cell(rr, 3, VALUATION[key])
            c.fill = INPUT
            c.font = BLUE
        vrows[key] = rr
        rr += 1
    wa.column_dimensions["A"].width = 46
    wa.column_dimensions["B"].width = 10
    for col in "CDE":
        wa.column_dimensions[col].width = 11

    # ------------------------------------------------------------ Historical
    wh = wb.create_sheet("Historical")
    wh["A1"] = "Reported figures, standalone bank (Rs bn unless stated)"
    wh["A1"].font = Font(bold=True, size=13)
    for j, h in enumerate(["Metric"] + YEARS_A + ["Source"]):
        c = wh.cell(3, 1 + j, h)
        c.font = BOLD
        c.fill = HEAD
    hrow = {}
    for k, (key, label) in enumerate(HIST_METRICS):
        rr = 4 + k
        wh.cell(rr, 1, label)
        for i, v in enumerate(hist[key]):
            if v is not None:
                c = wh.cell(rr, 2 + i, v)
                c.font = BLUE
                c.number_format = BN
        wh.cell(rr, 2 + n_a, src[key])
        hrow[key] = rr
    wh.column_dimensions["A"].width = 36
    wh.column_dimensions[get_column_letter(2 + n_a)].width = 60

    # ------------------------------------------------------------ Model
    wm = wb.create_sheet("Model")
    wm["A1"] = "Forecast model (Rs bn unless stated). A = reported, E = forecast."
    wm["A1"].font = Font(bold=True, size=13)
    cols = YEARS_A + YEARS_E
    for j, h in enumerate(["Line item"] + cols):
        c = wm.cell(3, 1 + j, h)
        c.font = BOLD
        c.fill = HEAD
    L = lambda i: get_column_letter(2 + i)  # column of year index i
    H = lambda key, i: f"Historical!{L(i)}{hrow[key]}"
    A = lambda key, t: f"Assumptions!{get_column_letter(3 + t)}{active[key]}"
    lines = ["Advances", "Advances growth (%)", "Total assets", "Average total assets",
             "Interest-earning assets / avg assets", "Average interest-earning assets",
             "Net interest margin (%)", "Net interest income", "NII growth (%)",
             "Non-interest income ex treasury", "Treasury income", "Core income",
             "Operating expenses", "Cost-to-income, core (%)", "Core operating profit",
             "Provisions", "Credit cost (%)", "Profit before tax", "Tax", "Tax rate (%)",
             "Profit after tax", "PAT growth (%)", "Net worth (capital + reserves)",
             "Average net worth", "ROE (%)", "ROA (%)", "Shares (bn)", "EPS (Rs)",
             "EPS growth (%)", "Book value per share (Rs)", "Dividend per share (Rs)",
             "Check: PBT vs reported (Rs bn)"]
    R = {name: 4 + k for k, name in enumerate(lines)}
    for name, rr in R.items():
        wm.cell(rr, 1, name)
    ref = lambda name, i: f"{L(i)}{R[name]}"
    for i in range(n_a + 3):
        t = i - n_a
        fc = t >= 0
        f = {}
        f["Advances"] = f"={ref('Advances', i - 1)}*(1+{A('loan_growth', t)}/100)" if fc else f"={H('total_advances_net', i)}"
        f["Total assets"] = f"={ref('Total assets', i - 1)}*(1+{A('loan_growth', t)}/100)" if fc else f"={H('total_assets', i)}"
        f["Net interest margin (%)"] = f"={A('nim', t)}" if fc else f"={H('net_interest_margin', i)}"
        if i >= 1:
            f["Advances growth (%)"] = f"=({ref('Advances', i)}/{ref('Advances', i - 1)}-1)*100"
            f["Average total assets"] = f"=({ref('Total assets', i - 1)}+{ref('Total assets', i)})/2"
            f["NII growth (%)"] = f"=({ref('Net interest income', i)}/{ref('Net interest income', i - 1)}-1)*100"
            f["PAT growth (%)"] = f"=({ref('Profit after tax', i)}/{ref('Profit after tax', i - 1)}-1)*100"
            f["EPS growth (%)"] = f"=({ref('EPS (Rs)', i)}/{ref('EPS (Rs)', i - 1)}-1)*100"
            f["Average net worth"] = f"=({ref('Net worth (capital + reserves)', i - 1)}+{ref('Net worth (capital + reserves)', i)})/2"
            f["ROE (%)"] = f"={ref('Profit after tax', i)}/{ref('Average net worth', i)}*100"
            f["ROA (%)"] = f"={ref('Profit after tax', i)}/{ref('Average total assets', i)}*100"
            f["Credit cost (%)"] = f"={ref('Provisions', i)}/(({ref('Advances', i - 1)}+{ref('Advances', i)})/2)*100"
            if fc:
                f["Interest-earning assets / avg assets"] = f"={ref('Interest-earning assets / avg assets', n_a - 1)}"
                f["Average interest-earning assets"] = f"={ref('Interest-earning assets / avg assets', i)}*{ref('Average total assets', i)}"
                f["Net interest income"] = f"={ref('Net interest margin (%)', i)}/100*{ref('Average interest-earning assets', i)}"
            else:
                f["Average interest-earning assets"] = f"={ref('Net interest income', i)}/({ref('Net interest margin (%)', i)}/100)"
                f["Interest-earning assets / avg assets"] = f"={ref('Average interest-earning assets', i)}/{ref('Average total assets', i)}"
        if not fc:
            f["Net interest income"] = f"={H('net_interest_income', i)}"
            f["Non-interest income ex treasury"] = f"={H('non_interest_income_excl_treasury', i)}"
            f["Treasury income"] = f"={H('treasury_income', i)}"
            f["Operating expenses"] = f"={H('operating_expenses', i)}"
            f["Provisions"] = f"={H('provisions', i)}"
            f["Tax"] = f"={H('tax', i)}"
            f["Net worth (capital + reserves)"] = f"={H('net_worth', i)}"
            f["Shares (bn)"] = f"={H('equity_share_capital', i)}/{FACE_VALUE}"
            f["Dividend per share (Rs)"] = f"={H('dividend_per_share', i)}"
            f["Check: PBT vs reported (Rs bn)"] = f"={ref('Profit before tax', i)}-{H('profit_before_tax', i)}"
        else:
            f["Non-interest income ex treasury"] = f"={ref('Non-interest income ex treasury', i - 1)}*(1+{A('nonii_growth', t)}/100)"
            f["Treasury income"] = f"={A('treasury', t)}"
            f["Operating expenses"] = f"={A('cost_income', t)}/100*{ref('Core income', i)}"
            f["Provisions"] = f"={A('credit_cost', t)}/100*({ref('Advances', i - 1)}+{ref('Advances', i)})/2"
            f["Tax"] = f"={A('tax_rate', t)}/100*{ref('Profit before tax', i)}"
            f["Net worth (capital + reserves)"] = (f"={ref('Net worth (capital + reserves)', i - 1)}+{ref('Profit after tax', i)}"
                                                   f"-{ref('Dividend per share (Rs)', i - 1)}*{ref('Shares (bn)', i - 1)}")
            f["Shares (bn)"] = f"={ref('Shares (bn)', n_a - 1)}"
            f["Dividend per share (Rs)"] = f"={A('payout', t)}/100*{ref('EPS (Rs)', i)}"
        f["Core income"] = f"={ref('Net interest income', i)}+{ref('Non-interest income ex treasury', i)}"
        f["Cost-to-income, core (%)"] = f"={ref('Operating expenses', i)}/{ref('Core income', i)}*100"
        f["Core operating profit"] = f"={ref('Core income', i)}-{ref('Operating expenses', i)}"
        f["Profit before tax"] = f"={ref('Core operating profit', i)}-{ref('Provisions', i)}+{ref('Treasury income', i)}"
        f["Tax rate (%)"] = f"={ref('Tax', i)}/{ref('Profit before tax', i)}*100"
        f["Profit after tax"] = f"={ref('Profit before tax', i)}-{ref('Tax', i)}"
        f["EPS (Rs)"] = f"={ref('Profit after tax', i)}/{ref('Shares (bn)', i)}"
        f["Book value per share (Rs)"] = f"={ref('Net worth (capital + reserves)', i)}/{ref('Shares (bn)', i)}"
        for name, formula in f.items():
            c = wm.cell(R[name], 2 + i, formula)
            c.number_format = PCT if "(%)" in name or "/ avg" in name else BN
    for rr in [R["Profit after tax"], R["ROE (%)"], R["EPS (Rs)"], R["Book value per share (Rs)"]]:
        wm.cell(rr, 1).font = BOLD
    wm.column_dimensions["A"].width = 38
    for i in range(n_a + 3):
        wm.column_dimensions[L(i)].width = 11
    wm.freeze_panes = "B4"

    # ------------------------------------------------------------ Valuation
    wv = wb.create_sheet("Valuation")
    wv["A1"] = "Valuation: residual income (primary) and justified P/B (cross-check)"
    wv["A1"].font = Font(bold=True, size=13)
    V = lambda key: f"Assumptions!$C${vrows[key]}"
    M = lambda name, i: f"Model!{L(i)}{R[name]}"
    wv["A3"], wv["B3"] = "Cost of equity (%)", f"={V('coe')}"
    wv["A4"], wv["B4"] = "Long-run growth (%)", f"={V('g')}"
    for j, h in enumerate(["", ] + YEARS_E):
        c = wv.cell(6, 1 + j, h)
        c.font = BOLD
        c.fill = HEAD
    wv["A7"], wv["A8"], wv["A9"], wv["A10"] = "Opening net worth", "Profit after tax", "Equity charge (COE x opening NW)", "Residual income"
    wv["A11"] = "Discount factor"
    wv["A12"] = "PV of residual income"
    for t in range(3):
        i = n_a + t
        col = get_column_letter(2 + t)
        wv[f"{col}7"] = f"={M('Net worth (capital + reserves)', i - 1)}"
        wv[f"{col}8"] = f"={M('Profit after tax', i)}"
        wv[f"{col}9"] = f"=$B$3/100*{col}7"
        wv[f"{col}10"] = f"={col}8-{col}9"
        wv[f"{col}11"] = f"=1/(1+$B$3/100)^{t + 1}"
        wv[f"{col}12"] = f"={col}10*{col}11"
    rows = [
        ("Net worth, 31 Mar 2026", f"={M('Net worth (capital + reserves)', n_a - 1)}"),
        ("Sum of PV of residual income FY27E-FY29E", "=SUM(B12:D12)"),
        ("Terminal value of residual income at FY29E", "=D10*(1+B4/100)/(B3/100-B4/100)"),
        ("PV of terminal value", "=B16*D11"),
        ("Equity value, 31 Mar 2026", "=B14+B15+B17"),
        ("Shares (bn)", f"={M('Shares (bn)', n_a - 1)}"),
        ("Value per share, 31 Mar 2026 (Rs)", "=B18/B19"),
        ("Value per share today (Rs)", f"=B20*(1+B3/100)^({V('months_since_fy26')}/12)-{V('dps_paid_since_fy26')}"),
        ("Standalone 12-month value (Rs)", f"=B21*(1+B3/100)-{M('Dividend per share (Rs)', n_a)}"),
        ("Listed subsidiaries after holdco discount (Rs)", f"=SOTP_PER_SHARE*(1-{V('holdco_discount')}/100)"),
        ("12-month target price (Rs)", "=B22+B23"),
        ("Current price (Rs)", f"={V('price')}"),
        ("Upside / downside (%)", "=(B24/B25-1)*100"),
        ("Rating (BUY > 15%, SELL < -5%, else HOLD)", '=IF(B26>15,"BUY",IF(B26<-5,"SELL","HOLD"))'),
        ("", None),
        ("Cross-check: justified P/B", None),
        ("Sustainable ROE, avg FY28E-FY29E (%)", f"=({M('ROE (%)', n_a + 1)}+{M('ROE (%)', n_a + 2)})/2"),
        ("Justified P/B = (ROE - g) / (COE - g)", "=(B30-B4)/(B3-B4)"),
        ("Book value per share in 12 months (Rs)",
         f"={M('Book value per share (Rs)', n_a)}+{V('months_since_fy26')}/12*({M('Book value per share (Rs)', n_a + 1)}-{M('Book value per share (Rs)', n_a)})"),
        ("P/B-based 12-month value incl. subsidiaries (Rs)", "=B31*B32+B23"),
        ("", None),
        ("Market multiples at current price", None),
        ("P/E FY27E", f"=B25/{M('EPS (Rs)', n_a)}"),
        ("P/E FY28E", f"=B25/{M('EPS (Rs)', n_a + 1)}"),
        ("P/B FY26A", f"=B25/{M('Book value per share (Rs)', n_a - 1)}"),
        ("P/B FY27E", f"=B25/{M('Book value per share (Rs)', n_a)}"),
        ("P/B FY28E", f"=B25/{M('Book value per share (Rs)', n_a + 1)}"),
    ]
    for k, (label, formula) in enumerate(rows):
        rr = 14 + k
        wv.cell(rr, 1, label)
        if formula:
            c = wv.cell(rr, 2, formula)
            c.number_format = "0.00"
        if label.startswith(("12-month target", "Rating", "Cross-check", "Market multiples")):
            wv.cell(rr, 1).font = BOLD
    # sensitivity: standalone 12m value + subs for COE x g
    s0 = 14 + len(rows) + 2
    wv.cell(s0, 1, "Sensitivity: 12-month target (Rs) by cost of equity (rows) and growth (columns)").font = BOLD
    gs = [7.0, 7.5, 8.0, 8.5, 9.0]
    cs = [11.0, 11.5, 12.0, 12.5, 13.0]
    for j, g in enumerate(gs):
        c = wv.cell(s0 + 1, 2 + j, g)
        c.font = BOLD
        c.fill = HEAD
    nw0 = M('Net worth (capital + reserves)', n_a - 1)
    sh = M('Shares (bn)', n_a - 1)
    for k, cval in enumerate(cs):
        rr = s0 + 2 + k
        cc = wv.cell(rr, 1, cval)
        cc.font = BOLD
        cc.fill = HEAD
        for j in range(len(gs)):
            c = f"($A{rr}/100)"
            g = f"({get_column_letter(2 + j)}${s0 + 1}/100)"
            ri = [f"({M('Profit after tax', n_a + t)}-{c}*{M('Net worth (capital + reserves)', n_a + t - 1)})" for t in range(3)]
            pv = "+".join(f"{ri[t]}/(1+{c})^{t + 1}" for t in range(3))
            tv = f"{ri[2]}*(1+{g})/({c}-{g})/(1+{c})^3"
            v_mar = f"(({nw0}+{pv}+{tv})/{sh})"
            v_today = f"({v_mar}*(1+{c})^({V('months_since_fy26')}/12)-{V('dps_paid_since_fy26')})"
            formula = f"={v_today}*(1+{c})-{M('Dividend per share (Rs)', n_a)}+$B$23"
            cell = wv.cell(rr, 2 + j, formula)
            cell.number_format = "#,##0"
    # sum of the parts: listed subsidiaries at market value
    t0 = s0 + 2 + len(cs) + 2
    wv.cell(t0, 1, "Sum of the parts: listed subsidiaries at market value").font = BOLD
    for j, h in enumerate(["Subsidiary", "Shares held (m)", "Price (Rs)", "Value (Rs bn)", "Source"]):
        c = wv.cell(t0 + 1, 1 + j, h)
        c.font = BOLD
        c.fill = HEAD
    for k, (name, shm, px, srcs) in enumerate(LISTED_SUBS):
        rr = t0 + 2 + k
        wv.cell(rr, 1, name)
        for col, v in ((2, shm), (3, px)):
            c = wv.cell(rr, col, v)
            c.fill = INPUT
            c.font = BLUE
        wv.cell(rr, 4, f"=B{rr}*C{rr}/1000").number_format = BN
        wv.cell(rr, 5, srcs)
    tot = t0 + 2 + len(LISTED_SUBS)
    wv.cell(tot, 1, "Total (Rs bn)").font = BOLD
    wv.cell(tot, 4, f"=SUM(D{t0 + 2}:D{tot - 1})").number_format = BN
    wv.cell(tot + 1, 1, "Per ICICI Bank share, before discount (Rs)").font = BOLD
    wv.cell(tot + 1, 4, f"=D{tot}/{sh}").number_format = "0.0"
    wv.cell(tot + 2, 1, "Unlisted subsidiaries (ICICI Securities, UK, Canada, Home Finance) are not valued.")
    for row in wv.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "SOTP_PER_SHARE" in c.value:
                c.value = c.value.replace("SOTP_PER_SHARE", f"$D${tot + 1}")
    wv.column_dimensions["A"].width = 52
    for col in "BCDEF":
        wv.column_dimensions[col].width = 12

    wb.save(path)
    return dict(model_rows=R, active=active, vrows=vrows)


def main():
    hist_path = os.path.join(ROOT, "data", "key_figures.csv")
    hist, src = load_history(hist_path)
    results = {}
    for name, drv in SCENARIOS.items():
        results[name] = run_model(hist, drv, VALUATION)
    out = dict(years=YEARS_A + YEARS_E, scenarios=results, valuation_inputs=VALUATION)
    with open(os.path.join(ROOT, "model", "model_output.json"), "w") as f:
        json.dump(out, f, indent=1)
    layout = write_excel(hist, src, os.path.join(ROOT, "model", "ICICI_Bank_Model.xlsx"))
    with open(os.path.join(ROOT, "model", "excel_layout.json"), "w") as f:
        json.dump(layout, f, indent=1)
    b = results["Base"]["valuation"]
    print(f"Base: target Rs {b['target']:.0f}, price {b['price']}, upside {b['upside']:.1f}% -> {b['rating']}")
    for s in ["Bull", "Bear"]:
        print(f"{s}: target Rs {results[s]['valuation']['target']:.0f}")


if __name__ == "__main__":
    main()
