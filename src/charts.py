"""Charts for the initiation report, from model/model_output.json and data/*.csv."""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "charts")
o = json.load(open(os.path.join(ROOT, "model", "model_output.json")))
yrs = o["years"]
b = o["scenarios"]["Base"]
NAVY, TEAL, GREY, ORANGE = "#1f3a5f", "#2a9d8f", "#9aa5b1", "#e76f51"
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
est = [y.endswith("E") for y in yrs]


def save(fig, name):
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, name), dpi=200)
    plt.close(fig)


# 1. NII and PAT
fig, ax = plt.subplots(figsize=(5.2, 2.6))
x = range(len(yrs))
ax.bar([i - 0.2 for i in x], b["nii"], 0.4, color=[GREY if e else NAVY for e in est], label="Net interest income")
ax.bar([i + 0.2 for i in x], b["pat"], 0.4, color=[GREY if e else TEAL for e in est], label="Profit after tax")
ax.set_xticks(list(x), yrs)
ax.set_ylabel("Rs bn")
ax.legend(frameon=False, loc="upper left")
ax.set_title("Earnings: NII and profit (grey = base-case forecast)", fontsize=9)
save(fig, "earnings.png")

# 2. NIM and ROE
fig, ax = plt.subplots(figsize=(5.2, 2.6))
ax.plot(yrs, b["nim"], marker="o", color=NAVY, label="NIM (%)")
ax2 = ax.twinx()
ax2.plot(yrs[1:], b["roe"][1:], marker="s", color=ORANGE, label="ROE (%)")
ax.axvspan(4.5, 7.5, color="#eeeeee", zorder=0)
ax.set_ylabel("NIM (%)")
ax2.set_ylabel("ROE (%)")
ax2.spines["right"].set_visible(True)
h1, l1 = ax.get_legend_handles_labels()
h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, frameon=False, loc="lower left")
ax.set_title("Margin and return on equity (shaded = forecast)", fontsize=9)
save(fig, "nim_roe.png")

# 3. Asset quality
hist = pd.read_csv(os.path.join(ROOT, "data", "key_figures.csv")).set_index("metric")
cols = ["fy2022", "fy2023", "fy2024", "fy2025", "fy2026"]
fig, ax = plt.subplots(figsize=(5.2, 2.4))
ax.plot(yrs[:5], hist.loc["gross_npa_ratio", cols].astype(float), marker="o", color=NAVY, label="Gross NPA (%)")
ax.plot(yrs[:5], hist.loc["net_npa_ratio", cols].astype(float), marker="o", color=TEAL, label="Net NPA (%)")
ax3 = ax.twinx()
ax3.bar(yrs[:5], hist.loc["credit_cost_provisions_to_avg_advances", cols].astype(float), color=GREY, alpha=0.5, label="Credit cost (%)", width=0.4)
ax3.spines["right"].set_visible(True)
ax3.set_ylabel("Credit cost (%)")
ax.set_zorder(ax3.get_zorder() + 1)
ax.patch.set_visible(False)
h1, l1 = ax.get_legend_handles_labels()
h2, l2 = ax3.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, frameon=False, loc="upper right")
ax.set_title("Asset quality", fontsize=9)
save(fig, "asset_quality.png")

# 4. Share price vs peers, rebased
names = {"icicibank": "ICICI Bank", "hdfcbank": "HDFC Bank", "axisbank": "Axis Bank",
         "kotakbank": "Kotak Mahindra", "sbin": "SBI"}
fig, ax = plt.subplots(figsize=(5.2, 2.6))
perf = {}
for f, n in names.items():
    d = pd.read_csv(os.path.join(ROOT, "data", f + ".csv"), parse_dates=["Date"])
    d = d[d["Date"] >= "2021-04-01"].set_index("Date")["Close"]
    perf[n] = d
    ax.plot(d.index, d / d.iloc[0] * 100, lw=1.6 if n == "ICICI Bank" else 0.9,
            color=NAVY if n == "ICICI Bank" else None, label=n)
ax.set_ylabel("Rebased, 1 Apr 2021 = 100")
ax.legend(frameon=False, fontsize=7, ncol=2)
ax.set_title("Share price vs large private and public peers", fontsize=9)
save(fig, "price_vs_peers.png")

# 5. Price-to-book band, using book value interpolated between year ends
d = perf["ICICI Bank"]
ends = pd.to_datetime(["2022-03-31", "2023-03-31", "2024-03-31", "2025-03-31", "2026-03-31", "2027-03-31"])
bv = pd.Series(b["bvps"][:6], index=ends)
bv_daily = bv.reindex(bv.index.union(d.index)).interpolate("time").reindex(d.index)
pb = (d / bv_daily).dropna()
pb = pb[pb.index >= "2022-03-31"]
fig, ax = plt.subplots(figsize=(5.2, 2.4))
ax.plot(pb.index, pb, color=NAVY, lw=1)
ax.axhline(pb.mean(), color=ORANGE, ls="--", lw=1, label=f"Average {pb.mean():.2f}x")
ax.axhline(o["scenarios"]["Base"]["valuation"]["justified_pb"], color=TEAL, ls=":", lw=1.2,
           label=f"Justified (base) {o['scenarios']['Base']['valuation']['justified_pb']:.2f}x")
ax.set_ylabel("Price / book (x)")
ax.legend(frameon=False, fontsize=7)
ax.set_title("Price-to-book on standalone book value", fontsize=9)
save(fig, "pb_band.png")

stats = {"pb_mean": float(pb.mean()), "pb_min": float(pb.min()), "pb_max": float(pb.max()),
         "pb_last": float(pb.iloc[-1]), "last_date": str(d.index[-1].date())}
for n, s in perf.items():
    last = s.index[-1]
    for lab, days in [("1y", 365), ("3y", 365 * 3)]:
        start = s[s.index <= last - pd.Timedelta(days=days)].iloc[-1]
        stats[f"{n}|{lab}"] = float((s.iloc[-1] / start - 1) * 100)
json.dump(stats, open(os.path.join(ROOT, "model", "chart_stats.json"), "w"), indent=1)
print(json.dumps(stats, indent=1))
