"""Extract ICICI Bank's standalone figures from its results filings.

Inputs  raw/pdf/*.pdf   Q4 investor presentations FY22-FY26, the Q2 FY27
                        presentation, and audited results / press releases.
Outputs data/key_figures.csv     one row per metric, FY22-FY26 plus H1 FY27,
                                 each with the file and page it came from
        data/reconciliation.csv  every cross-check that was run, with the result

How it works: each PDF is converted with `pdftotext -layout`, which keeps a
slide's table rows on one line. For each metric we find the slide by its title,
the row by its label, and take the column for the year. Footnote markers that
pdftotext glues onto numbers (53.80 + footnote 1 -> "53.801") are removed by
rounding to the number of decimals the bank reports for that row.

Checks:
  1. Prior-year columns: each presentation also prints last year's figure.
     It must equal what last year's presentation reported (catches restatements
     and parsing errors).
  2. Audited results: profit after tax, deposits, advances, total assets and
     share capital (in Rs crore) must match the presentation (in Rs billion).
  3. Arithmetic: core operating profit - provisions + treasury = PBT, and
     PBT - tax = PAT, for every year.
Exits non-zero if any check fails.
"""
import csv
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "raw", "pdf")

YEARS = ["fy2022", "fy2023", "fy2024", "fy2025", "fy2026"]
DECKS = {y: f"fy{y[-2:]}_q4_investor_presentation.pdf" for y in YEARS}
H1_DECK = "fy27_q2_investor_presentation.pdf"

# metric: (slide title regex, row label regex, decimals, unit, uses % tokens, column kind)
# column kind: "pl" = [FY prev, Q4 prev, Q3, Q4, FY cur] ; "bs" = [prev, mid, cur] ; "pct_last"
METRICS = {
    "net_interest_income": (r"Profit & loss statement", r"Net interest income", 2, "Rs bn", False, "pl"),
    "non_interest_income_excl_treasury": (r"Profit & loss statement", r"Non-interest income", 2, "Rs bn", False, "pl"),
    "fee_income": (r"Profit & loss statement", r"- Fee income", 2, "Rs bn", False, "pl"),
    "operating_expenses": (r"Profit & loss statement", r"Operating expenses", 2, "Rs bn", False, "pl"),
    "core_operating_profit": (r"Profit & loss statement", r"Core operating profit(?! *excluding)(?!\s*$)", 2, "Rs bn", False, "pl"),
    "provisions": (r"Profit & loss statement", r"(Net provisions|Provisions)", 2, "Rs bn", False, "pl"),
    "treasury_income": (r"Profit & loss statement", r"Treasury income", 2, "Rs bn", False, "pl"),
    "profit_before_tax": (r"Profit & loss statement", r"Profit before tax(?! *excluding)(?!\s*$)", 2, "Rs bn", False, "pl"),
    "tax": (r"Profit & loss statement", r"Tax", 2, "Rs bn", False, "pl"),
    "profit_after_tax": (r"Profit & loss statement", r"Profit after tax", 2, "Rs bn", False, "pl"),
    "net_interest_margin": (r"Key ratios$", r"Net interest margin", 2, "%", False, "pl"),
    "cost_to_income": (r"Key ratios$", r"Cost-to-income", 1, "%", False, "pl"),
    "credit_cost_provisions_to_avg_advances": (r"Key ratios$", r"Provisions/average advances", 2, "%", False, "pl"),
    "roa": (r"Key ratios$", r"Return on average assets", 2, "%", False, "pl"),
    "eps_basic": (r"Key ratios$", r"Weighted average EPS", 1, "Rs", False, "pl"),
    "book_value_per_share": (r"Key ratios$", r"Book value", 1, "Rs", False, "pl"),
    "net_worth": (r"Balance sheet: liabilities", r"Net worth", 2, "Rs bn", False, "bs"),
    "equity_share_capital": (r"Balance sheet: liabilities", r"- Equity capital", 2, "Rs bn", False, "bs"),
    "total_deposits": (r"Balance sheet: liabilities", r"Deposits", 2, "Rs bn", False, "bs"),
    "current_deposits": (r"Balance sheet: liabilities", r"- Current", 2, "Rs bn", False, "bs"),
    "savings_deposits": (r"Balance sheet: liabilities", r"- Savings", 2, "Rs bn", False, "bs"),
    "total_assets": (r"Balance sheet: liabilities", r"Total liabilities", 2, "Rs bn", False, "bs"),
    "total_advances_net": (r"Balance sheet: assets", r"Advances", 2, "Rs bn", False, "bs"),
    "gross_npa_amount": (r"NPA trends", r"Gross NPAs", 2, "Rs bn", False, "bs"),
    "net_npa_amount": (r"NPA trends", r"Net NPAs", 2, "Rs bn", False, "bs"),
    "gross_npa_ratio": (r"NPA trends", r"Gross NPA ratio", 2, "%", True, "bs"),
    "net_npa_ratio": (r"NPA trends", r"Net NPA ratio", 2, "%", True, "bs"),
    "provision_coverage_ratio": (r"NPA trends", r"Provision coverage ratio", 1, "%", True, "bs"),
    "cet1_ratio": (r"(Standalone capital adequacy|Strong capital position)", r"- of which: CET1", 2, "%", None, "pct_last"),
    "total_car": (r"(Standalone capital adequacy|Strong capital position)", r"Total capital", 2, "%", None, "pct_last"),
}

NUM = re.compile(r"^\(?-?[\d,]*\d(\.\d+)?\)?\d*%?$")


def pages(pdf):
    out = subprocess.run(["pdftotext", "-layout", os.path.join(PDF, pdf), "-"],
                         capture_output=True, text=True, check=True).stdout
    return out.split("\f")


def find_slide(pg, title):
    """First page whose first non-empty line matches the title (standalone, not consolidated)."""
    for i, p in enumerate(pg):
        lines = [l.strip() for l in p.splitlines() if l.strip()]
        if lines and re.match(title, lines[0]) and "onsolidated" not in lines[0]:
            return i
    return None


def parse_num(tok, dec, pct):
    tok = re.sub(r"\)\d+$", ")", tok)      # "(2.81)1": footnote after a negative
    if pct is None:                        # either form accepted
        tok = tok.rstrip("%")
    elif tok.endswith("%"):
        if not pct:
            return None
        tok = tok[:-1]
    elif pct:
        return None
    neg = tok.startswith("(") and tok.endswith(")")
    tok = tok.strip("()").replace(",", "")
    if "." not in tok:
        return None
    whole, frac = tok.split(".")
    if len(frac) > dec:          # footnote digit glued on: drop it
        frac = frac[:dec]
    v = float(f"{whole}.{frac}")
    return -v if neg else v


def row_values(page, label, dec, pct):
    """Numbers on the first line whose text starts with the label."""
    for line in page.splitlines():
        s = line.strip()
        m = re.match(rf"{label}(\s*\d(,\d)*)?\s", s + " ")
        if not m:
            continue
        rest = s[m.end():]
        vals = []
        for tok in rest.split():
            if NUM.match(tok):
                v = parse_num(tok, dec, pct)
                if v is not None:
                    vals.append(v)
        if vals:
            return vals
    return None


def get(pg, metric, col):
    title, label, dec, unit, pct, kind = METRICS[metric]
    # P&L metrics live on two consecutive slides with the same title
    start = find_slide(pg, title)
    if start is None:
        return None, None
    for i in range(start, min(start + 3, len(pg))):
        first = next((l.strip() for l in pg[i].splitlines() if l.strip()), "")
        if not re.match(title, first):
            continue
        vals = row_values(pg[i], label, dec, pct)
        if not vals:
            continue
        if kind == "pct_last":
            idx = {"cur": -1, "prev": 0}[col]
        elif kind == "bs":
            idx = {"cur": 2, "prev": 0}[col]
        else:
            idx = {"cur": 4, "prev": 0, "h1": 5, "q2": 4}[col]
        if len(vals) <= max(idx, -idx - 1 if idx < 0 else idx):
            continue
        return vals[idx], i + 1
    return None, None


def audited(pdf, kind):
    """Rs crore figures from an audited results PDF / press release, converted to Rs bn.
    Returns {metric: (current, prior)}; annual P&L columns are [Q4, Q3, Q4 prev, FY, FY prev]."""
    txt = "\f".join(pages(pdf))
    out = {}
    lines = txt.splitlines()

    def nums(s):
        return [float(t.replace(",", "")) for t in s.split() if re.match(r"^[\d,]+\.\d\d$", t)]
    for i, l in enumerate(lines):
        s = l.strip()
        if s.upper().startswith("15.") and "NET PROFIT FOR THE PERIOD" in s.upper() and "pat" not in out:
            v = nums(s) or nums(lines[i + 1])
            if kind == "annual" and len(v) >= 5:
                out["profit_after_tax"] = (v[3] / 100, v[4] / 100)
        for key, lab in [("total_deposits", "Deposits"), ("total_advances_net", "Advances"),
                         ("total_assets", "TOTAL CAPITAL AND LIABILITIES"), ("equity_share_capital", "Capital")]:
            if key not in out and re.match(rf"^{lab}\s+[\d,]+\.\d\d", s, re.I):
                v = nums(s)
                out[key] = (v[0] / 100, v[-1] / 100 if kind == "annual" else None)
    return out


def main():
    data, src, prev_seen = {}, {}, {}
    for y in YEARS:
        pg = pages(DECKS[y])
        for m in METRICS:
            v, p = get(pg, m, "cur")
            data.setdefault(m, {})[y] = v
            src.setdefault(m, {})[y] = f"{DECKS[y]} p{p}" if p else ""
            pv, _ = get(pg, m, "prev")
            prev_seen.setdefault(m, {})[y] = pv
    pg = pages(H1_DECK)
    h1_kinds = {"pl": "h1", "bs": "cur", "pct_last": "cur"}
    for m in METRICS:
        v, p = get(pg, m, h1_kinds[METRICS[m][5]])
        data[m]["h1fy2027"] = v
        src[m]["h1fy2027"] = f"{H1_DECK} p{p}" if p else ""

    # dividend per share, from the "dividend of Rs X per share" sentence in each deck
    data["dividend_per_share"], src["dividend_per_share"] = {}, {}
    for y in YEARS:
        pg = pages(DECKS[y])
        for i, p in enumerate(pg):
            mm = re.search(r"dividend of\s*[`₹]\s*(\d+(\.\d+)?) per (equity )?share", p)
            if mm:
                data["dividend_per_share"][y] = float(mm.group(1))
                src["dividend_per_share"][y] = f"{DECKS[y]} p{i + 1}"
                break

    # derived, clearly labelled
    data["casa_ratio_period_end"] = {}
    src["casa_ratio_period_end"] = {}
    for y in YEARS + ["h1fy2027"]:
        c, s_, d = (data["current_deposits"][y], data["savings_deposits"][y], data["total_deposits"][y])
        if None not in (c, s_, d):
            data["casa_ratio_period_end"][y] = round((c + s_) / d * 100, 1)
            src["casa_ratio_period_end"][y] = "derived: (current + savings) / deposits"

    checks = []

    def check(name, a, b, tol):
        ok = a is not None and b is not None and abs(a - b) <= tol
        checks.append((name, a, b, "OK" if ok else "FAIL"))

    # 1. prior-year column vs last year's deck
    for k in range(1, len(YEARS)):
        y, yp = YEARS[k], YEARS[k - 1]
        for m in METRICS:
            a, b = prev_seen[m][y], data[m][yp]
            if METRICS[m][5] == "pct_last":   # capital slides show Dec and Mar, not last March
                continue
            if a is None and b is None:
                continue
            tol = 0.011 if METRICS[m][3] == "Rs bn" else 0.011
            check(f"{m} {yp}: as printed in {y} deck vs {yp} deck", a, b, tol)
    # 2. audited results (Rs crore -> Rs bn)
    aud = {"fy2022": "fy22_q4_press_release.pdf", "fy2023": "fy23_q4_press_release.pdf",
           "fy2025": "fy25_q4_financial_results.pdf", "fy2026": "fy26_q4_financial_results.pdf"}
    for y, f in aud.items():
        a = audited(f, "annual")
        for m, (cur, prior) in a.items():
            check(f"{m} {y}: audited ({f}) vs deck", round(cur, 2), data[m][y], 0.011)
            yi = YEARS.index(y)
            if yi > 0 and prior is not None:
                check(f"{m} {YEARS[yi - 1]}: audited prior-year column ({f}) vs deck",
                      round(prior, 2), data[m][YEARS[yi - 1]], 0.011)
    a = audited("fy27_q2_financial_results.pdf", "half")
    for m, (cur, _) in a.items():
        check(f"{m} Sep 2026: audited (fy27_q2_financial_results.pdf) vs deck", round(cur, 2), data[m]["h1fy2027"], 0.011)
    # 3. arithmetic
    for y in YEARS + ["h1fy2027"]:
        d = {m: data[m][y] for m in METRICS}
        check(f"PBT {y}: core op profit - provisions + treasury", round(d["core_operating_profit"] - d["provisions"] + d["treasury_income"], 2), d["profit_before_tax"], 0.02)
        check(f"PAT {y}: PBT - tax", round(d["profit_before_tax"] - d["tax"], 2), d["profit_after_tax"], 0.02)
        check(f"Core op profit {y}: NII + non-interest income - opex", round(d["net_interest_income"] + d["non_interest_income_excl_treasury"] - d["operating_expenses"], 2), d["core_operating_profit"], 0.02)

    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    cols = YEARS + ["h1fy2027"]
    with open(os.path.join(ROOT, "data", "key_figures.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["metric"] + cols + ["unit"] + [f"source_{c}" for c in cols])
        for m in data:
            unit = METRICS[m][3] if m in METRICS else ("Rs" if m == "dividend_per_share" else "%")
            w.writerow([m] + ["" if data[m].get(c) is None else data[m][c] for c in cols] + [unit] +
                       [src[m].get(c, "") for c in cols])
    with open(os.path.join(ROOT, "data", "reconciliation.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["check", "value_a", "value_b", "result"])
        w.writerows(checks)
    missing = [(m, c) for m in data for c in cols if data[m].get(c) is None
               and not (m == "dividend_per_share" and c == "h1fy2027")]
    fails = [c for c in checks if c[3] != "OK"]
    print(f"{len(data)} metrics x {len(cols)} periods; missing {len(missing)}; checks {len(checks)}, failed {len(fails)}")
    for m in missing:
        print("  missing", *m)
    for c in fails:
        print("  FAIL", c)
    return 1 if fails or missing else 0


if __name__ == "__main__":
    sys.exit(main())
