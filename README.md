# ICICI Bank: equity research initiation

A sell-side style initiation on ICICI Bank (NSE: ICICIBANK). The work covers the full workflow: extracting numbers from the bank's filings, checking them, forecasting, valuing the bank and writing the report. Every number in the report traces back to a page in a filing.

**Rating: HOLD. 12-month target Rs 1,353 vs Rs 1,394 (23 Oct 2026).**
- Bull case: Rs 1,633.
- Bear case: Rs 891.

## Pipeline

```
raw/pdf/*.pdf ──extract.py──> data/key_figures.csv + data/reconciliation.csv
                                   │
               assumptions.py ──build_model.py──> model/ICICI_Bank_Model.xlsx + model_output.json
                                   │                     │
                              check_model.py        charts.py ──> charts/*.png
                                   │                     │
                                   └──── make_report.py ─┴──> report/ICICI_Bank_Initiation.pdf
```

| Step | What it does |
|---|---|
| `src/extract.py` | Reads 12 filings: Q4 investor presentations FY22-FY26, the Q2 FY27 presentation, and audited results / press releases. It pulls 32 standalone metrics for FY22-FY26 and H1 FY27, and records the file and page of each. It then runs **169 checks**, listed below. |
| `src/assumptions.py` | Holds the bull / base / bear forecast drivers and the valuation inputs. |
| `src/build_model.py` | Forecasts FY27E-FY29E and runs a residual income valuation, a justified P/B cross-check and a sum of the parts for the listed subsidiaries. It writes the Excel model with live formulas and a scenario switch. |
| `src/check_model.py` | Recalculates the workbook in LibreOffice. It confirms 0 formula errors and that Excel matches the Python model. |
| `src/charts.py`, `src/make_report.py` | Draw the charts and write the LaTeX report from the model output, so the text never drifts from the numbers. |

The 169 checks in `extract.py`:
- Each year's prior-year column matches last year's filing, which catches restatements and parsing errors.
- Profit, deposits, advances, assets and share capital match the audited results, which are in Rs crore.
- Core operating profit - provisions + treasury = PBT, and PBT - tax = PAT.

## Run

```
python3 src/extract.py && python3 src/build_model.py && python3 src/check_model.py \
  && python3 src/charts.py && python3 src/make_report.py \
  && cd report && pdflatex ICICI_Bank_Initiation.tex
```

Requirements: Python 3 with the packages in requirements.txt; poppler (`pdftotext`); LibreOffice (only for the check); and pdflatex.

## Data

- **Filings:** from ICICI Bank's investor relations site. Standalone figures, Rs billion.
- **Share prices:** NSE daily closes.
- **Subsidiary stakes:** September 2026 shareholding disclosures.
- **10-year G-sec yield:** 7.40% on 30 Oct 2026.

## Limitations

- The model covers the standalone bank. Listed subsidiaries enter at market value with a 20% holding-company discount, and unlisted ones are valued at zero.
- Forecasts are annual.
- This is a student research exercise, not investment advice.
