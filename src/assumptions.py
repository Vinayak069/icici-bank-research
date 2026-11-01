"""Forecast drivers and valuation inputs. Edit here, then rerun build_model.py.

Every driver is per year: FY27E, FY28E, FY29E. Percentages are written as
percent (15.0 means 15%).
"""

YEARS_E = ["FY27E", "FY28E", "FY29E"]

SCENARIOS = {
    "Base": {
        "loan_growth":  [22.0, 15.0, 14.0],   # H1 FY27 already +17.9% (FCNR(B)-backed loans); FY23-26: 18.7, 16.2, 13.3, 15.8
        "nim":          [4.20, 4.20, 4.25],   # FY26 4.32; H1 FY27 4.27, Q2 4.18
        "nonii_growth": [14.0, 13.0, 13.0],   # H1 FY27 +14.2% y-o-y
        "cost_income":  [38.5, 38.5, 38.5],   # core (ex treasury); FY26 40.2, H1 FY27 38.2
        "credit_cost":  [0.30, 0.45, 0.50],   # FY24-26 0.32-0.38; H1 FY27 0.25 annualised
        "tax_rate":     [24.5, 24.5, 24.5],   # FY24-26 ~24.3-25.0
        "payout":       [17.0, 17.0, 17.0],   # FY26 DPS 12 / EPS 70.2
        "treasury":     [0.0, 0.0, 0.0],      # Rs bn; not forecast
    },
    "Bull": {
        "loan_growth":  [25.0, 18.0, 16.0],
        "nim":          [4.27, 4.35, 4.40],
        "nonii_growth": [15.0, 15.0, 15.0],
        "cost_income":  [38.0, 37.5, 37.5],
        "credit_cost":  [0.25, 0.35, 0.40],
        "tax_rate":     [24.5, 24.5, 24.5],
        "payout":       [17.0, 17.0, 17.0],
        "treasury":     [0.0, 0.0, 0.0],
    },
    "Bear": {
        "loan_growth":  [19.0, 11.0, 11.0],
        "nim":          [4.10, 4.00, 4.00],
        "nonii_growth": [10.0, 9.0, 9.0],
        "cost_income":  [39.5, 40.5, 40.5],
        "credit_cost":  [0.40, 0.90, 0.80],
        "tax_rate":     [24.5, 24.5, 24.5],
        "payout":       [17.0, 17.0, 17.0],
        "treasury":     [0.0, 0.0, 0.0],
    },
}

DRIVER_LABELS = {
    "loan_growth": "Loan growth (%)",
    "nim": "Net interest margin (%)",
    "nonii_growth": "Non-interest income growth, ex treasury (%)",
    "cost_income": "Cost-to-income, core (%)",
    "credit_cost": "Credit cost, provisions / avg loans (%)",
    "tax_rate": "Tax rate (%)",
    "payout": "Dividend payout (%)",
    "treasury": "Treasury income (Rs bn)",
}

VALUATION = {
    "rf": 7.40,         # India 10Y G-sec yield, 30 Oct 2026 (Trading Economics)
    "beta": 1.0,
    "erp": 5.0,         # equity risk premium, %
    "g": 8.0,           # long-run growth of residual income / book, %
    "price": 1394.2,    # NSE close, Rs
    "price_date": "23 Oct 2026",
    "dps_paid_since_fy26": 12.0,   # FY26 final dividend, paid after March 2026
    "months_since_fy26": 7,        # 31 Mar 2026 -> 31 Oct 2026
    "holdco_discount": 20.0,       # % discount applied to subsidiary value
}

# Listed subsidiaries at market value (sum of the parts). Shares held by ICICI
# Bank, Sep 2026 shareholding; prices are NSE closes on 23 Oct 2026.
# Unlisted subsidiaries (ICICI Securities, ICICI Bank UK and Canada, ICICI Home
# Finance and others) are given no value, which is conservative.
LISTED_SUBS = [
    # name, shares held (million), price (Rs), source
    ("ICICI Life Insurance (ex ICICI Prudential Life)", 766.621197, 422.20,
     "52.83% stake, Trendlyne Sep-2026 shareholding; bank filing 2 Sep 2026"),
    ("ICICI Lombard General Insurance", 255.550678, 1592.00,
     "51.15% stake, company form NL-9, 30 Sep 2026"),
    ("ICICI Prudential AMC", 261.957214, 2651.80,
     "53.00% stake, Trendlyne Sep-2026 shareholding; listed 19 Dec 2025"),
]
