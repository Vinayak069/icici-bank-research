"""Recalculate the Excel model in LibreOffice and check it matches the Python model.

Usage: python3 src/check_model.py   (needs soffice on PATH)
Exits non-zero on any formula error or mismatch.
"""
import json
import os
import subprocess
import sys
import tempfile

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
xlsx = os.path.join(ROOT, "model", "ICICI_Bank_Model.xlsx")
out = json.load(open(os.path.join(ROOT, "model", "model_output.json")))
lay = json.load(open(os.path.join(ROOT, "model", "excel_layout.json")))
base = out["scenarios"]["Base"]

tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, xlsx],
               check=True, capture_output=True)
wb = load_workbook(os.path.join(tmp, "ICICI_Bank_Model.xlsx"), data_only=True)

errors = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#"):
                print("formula error", ws.title, c.coordinate, c.value)
                errors += 1

R = lay["model_rows"]
m = wb["Model"]
pairs = {"Profit after tax": "pat", "Net interest income": "nii", "ROE (%)": "roe",
         "EPS (Rs)": "eps", "Book value per share (Rs)": "bvps", "Provisions": "provisions",
         "Net worth (capital + reserves)": "net_worth", "Dividend per share (Rs)": "dps"}
bad = 0
for name, key in pairs.items():
    for i, v in enumerate(base[key]):
        if v is None:
            continue
        x = m.cell(R[name], 2 + i).value
        if x is None or abs(x - v) > 1e-6 * max(1, abs(v)):
            print("mismatch", name, out["years"][i], x, v)
            bad += 1
for i in range(5):
    chk = m.cell(R["Check: PBT vs reported (Rs bn)"], 2 + i).value
    if abs(chk) > 0.05:
        print("PBT does not reconcile", out["years"][i], chk)
        bad += 1
v = wb["Valuation"]
target = next(v.cell(r, 2).value for r in range(1, 80) if v.cell(r, 1).value == "12-month target price (Rs)")
if abs(target - base["valuation"]["target"]) > 1e-6 * target:
    print("target mismatch", target, base["valuation"]["target"])
    bad += 1
print(f"formula errors: {errors}, mismatches: {bad}, Excel target Rs {target:.2f}")
sys.exit(1 if errors or bad else 0)
