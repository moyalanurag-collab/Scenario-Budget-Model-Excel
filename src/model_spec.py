"""Single source of truth for the workbook: assumptions, scenario definitions and a Python mirror of the Excel engine.

The mirror is used for two things only:
  1. to fill the locked 'Budget' sheet with the Base case values at budget approval, and
  2. to independently verify the workbook's live Excel formulas after Excel recalculates them (see verify.ps1 / verify.py).

Everything is illustrative: the company, its numbers and the 'actuals' are fictional (see the README).
"""
import numpy as np

MONTHS = ["Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27"]
N_ACTUAL = 6   # Apr-Sep are actuals, Oct-Mar are forecast

# key, label, unit, base, upside, downside, number format, rationale
# 'shared' drivers have the same value in all three scenarios (Upside/Downside link to Base in the workbook).
ASSUMPTIONS = [
    ("section", "Customers and pricing"),
    ("open_customers", "Opening customers (1 Apr)", "customers", 420, None, None, "int", "Starting point of the financial year (illustrative)."),
    ("new_base", "New customers in April", "customers / month", 28, 34, 21, "int", "Sales pipeline run-rate; Upside assumes a stronger pipeline, Downside a slower sales cycle."),
    ("new_growth", "Monthly growth in new customers", "% per month", 0.020, 0.030, 0.005, "pct1", "Compounding monthly growth applied to April's new customers."),
    ("churn", "Monthly churn (share of opening customers)", "% per month", 0.018, 0.014, 0.026, "pct1", "About 19% / 15% / 27% annualised churn; typical range for small-business SaaS."),
    ("arpa", "Average revenue per customer per month (ARPA)", "GBP / month", 310, 322, 295, "gbp", "Subscription price mix; Upside includes more premium-plan customers."),
    ("price_inc", "Price increase", "%", 0.04, 0.05, 0.02, "pct1", "Annual list-price uplift applied to all customers from the month below."),
    ("price_month", "Price increase effective month (1 = Apr)", "month number", 7, None, None, "int", "Effective 1 October."),
    ("onboard_fee", "Onboarding fee per new customer", "GBP", 450, None, None, "gbp", "One-off implementation fee."),
    ("section", "Cost of sales"),
    ("host_cost", "Hosting cost per average customer", "GBP / month", 14, 14, 15, "gbp", "Cloud infrastructure; Downside assumes cost inflation."),
    ("support_cost", "Customer support cost per average customer", "GBP / month", 12, None, None, "gbp", "Support tooling and outsourced tier-1."),
    ("pay_fee", "Payment processing fees", "% of revenue", 0.015, None, None, "pct1", "Card / direct-debit fees."),
    ("section", "People and operating costs"),
    ("hc_open", "Opening headcount", "FTE", 24, None, None, "int", "Hiring plan is set out in the table below."),
    ("salary", "Average salary", "GBP / year", 52000, None, None, "gbp", "Blended across roles."),
    ("oncost", "Employer on-costs (NI, pension, benefits)", "% of salary", 0.18, None, None, "pct1", "Applied to salaries."),
    ("cac", "Marketing spend per new customer (CAC)", "GBP", 620, 580, 700, "gbp", "Acquisition marketing; Downside = less efficient channels."),
    ("tools_per_head", "Software and tools per head", "GBP / FTE / month", 180, None, None, "gbp", "Licences, laptops, collaboration tools."),
    ("rent", "Rent and facilities", "GBP / month", 9500, None, None, "gbp", "Office lease; fixed."),
    ("prof", "Professional fees", "GBP / month", 4200, None, None, "gbp", "Audit, legal, payroll bureau; fixed."),
    ("other", "Other G&A", "GBP / month", 6000, None, None, "gbp", "Insurance, travel, general admin; fixed."),
    ("section", "Capital, tax and cash"),
    ("capex_m", "Capital expenditure", "GBP / month", 8000, None, None, "gbp", "Equipment and capitalised development."),
    ("dep_life", "Depreciation life of new capex", "months", 36, None, None, "int", "Straight line, starting the month after purchase."),
    ("dep_existing", "Depreciation on existing assets", "GBP / month", 4500, None, None, "gbp", "Run-off of the opening asset base."),
    ("tax_rate", "Corporation tax rate", "%", 0.25, None, None, "pct1", "UK main rate; charged on positive monthly profit only (simplification: no loss relief)."),
    ("coll_in_month", "Share of revenue collected in the month billed", "%", 0.80, None, None, "pct1", "Remainder collected the following month."),
    ("open_cash", "Opening cash (1 Apr)", "GBP", 620000, None, None, "gbp", "Bank balance at start of year."),
    ("open_recv", "Opening trade receivables", "GBP", 24000, None, None, "gbp", "Collected in April."),
    ("open_tax", "Opening tax creditor", "GBP", 9000, None, None, "gbp", "Paid in April."),
]

# new hires starting in each month, by scenario (index 0 = Apr)
HIRES = {
    "Base":     [0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 1, 0],
    "Upside":   [0, 1, 1, 0, 1, 1, 0, 1, 1, 0, 1, 0],
    "Downside": [0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0],
}
SCENARIOS = ["Base", "Upside", "Downside"]


def drivers(scenario: str) -> dict:
    col = {"Base": 3, "Upside": 4, "Downside": 5}[scenario]
    out = {}
    for row in ASSUMPTIONS:
        if row[0] == "section":
            continue
        v = row[col]
        out[row[0]] = row[3] if v is None else v
    return out


def engine(d: dict, hires: list) -> dict:
    """Python mirror of the Excel engine sheets (12 monthly values per line)."""
    n = 12
    r = {k: [0.0] * n for k in [
        "open_cust", "new", "churned", "close_cust", "avg_cust", "arpa", "sub_rev", "onb_rev", "revenue", "hosting", "support", "payfees",
        "cogs", "gross_profit", "hc_open", "hires", "hc_close", "salaries", "marketing", "tools", "rent", "prof", "other", "opex", "ebitda",
        "da", "ebit", "tax", "net_income", "receipts", "cogs_paid", "opex_paid", "tax_paid", "capex", "net_cash", "open_cash", "close_cash",
        "mrr", "arr"]}
    for i in range(n):
        idx = i + 1
        r["open_cust"][i] = d["open_customers"] if i == 0 else r["close_cust"][i - 1]
        r["new"][i] = round_half_up(d["new_base"] * (1 + d["new_growth"]) ** (idx - 1))
        r["churned"][i] = round_half_up(r["open_cust"][i] * d["churn"])
        r["close_cust"][i] = r["open_cust"][i] + r["new"][i] - r["churned"][i]
        r["avg_cust"][i] = (r["open_cust"][i] + r["close_cust"][i]) / 2
        r["arpa"][i] = d["arpa"] * (1 + d["price_inc"] if idx >= d["price_month"] else 1)
        r["sub_rev"][i] = r["avg_cust"][i] * r["arpa"][i]
        r["onb_rev"][i] = r["new"][i] * d["onboard_fee"]
        r["revenue"][i] = r["sub_rev"][i] + r["onb_rev"][i]
        r["hosting"][i] = r["avg_cust"][i] * d["host_cost"]
        r["support"][i] = r["avg_cust"][i] * d["support_cost"]
        r["payfees"][i] = r["revenue"][i] * d["pay_fee"]
        r["cogs"][i] = r["hosting"][i] + r["support"][i] + r["payfees"][i]
        r["gross_profit"][i] = r["revenue"][i] - r["cogs"][i]
        r["hc_open"][i] = d["hc_open"] if i == 0 else r["hc_close"][i - 1]
        r["hires"][i] = hires[i]
        r["hc_close"][i] = r["hc_open"][i] + r["hires"][i]
        r["salaries"][i] = r["hc_close"][i] * d["salary"] / 12 * (1 + d["oncost"])
        r["marketing"][i] = r["new"][i] * d["cac"]
        r["tools"][i] = r["hc_close"][i] * d["tools_per_head"]
        r["rent"][i], r["prof"][i], r["other"][i] = d["rent"], d["prof"], d["other"]
        r["opex"][i] = r["salaries"][i] + r["marketing"][i] + r["tools"][i] + r["rent"][i] + r["prof"][i] + r["other"][i]
        r["ebitda"][i] = r["gross_profit"][i] - r["opex"][i]
        r["capex"][i] = d["capex_m"]
        r["da"][i] = d["dep_existing"] + (d["capex_m"] * i) / d["dep_life"]
        r["ebit"][i] = r["ebitda"][i] - r["da"][i]
        r["tax"][i] = max(0.0, r["ebit"][i]) * d["tax_rate"]
        r["net_income"][i] = r["ebit"][i] - r["tax"][i]
        prev_rev = None if i == 0 else r["revenue"][i - 1]
        r["receipts"][i] = r["revenue"][i] * d["coll_in_month"] + (d["open_recv"] if i == 0 else prev_rev * (1 - d["coll_in_month"]))
        r["cogs_paid"][i] = r["cogs"][i]
        r["opex_paid"][i] = r["opex"][i]
        r["tax_paid"][i] = d["open_tax"] if i == 0 else r["tax"][i - 1]
        r["net_cash"][i] = r["receipts"][i] - r["cogs_paid"][i] - r["opex_paid"][i] - r["tax_paid"][i] - r["capex"][i]
        r["open_cash"][i] = d["open_cash"] if i == 0 else r["close_cash"][i - 1]
        r["close_cash"][i] = r["open_cash"][i] + r["net_cash"][i]
        r["mrr"][i] = r["close_cust"][i] * r["arpa"][i]
        r["arr"][i] = r["mrr"][i] * 12
    return r


def round_half_up(x: float) -> int:
    """Excel's ROUND() rounds halves away from zero (Python's round() rounds half to even)."""
    return int(np.floor(x + 0.5))


def synthetic_actuals() -> dict:
    """Fictional 'actual' results for Apr-Sep: the same engine run on a plausible 'what really happened' set of drivers
    (slower new sales, higher churn and hosting cost, delayed hires, a legal-fee overrun) with +/-1.5% random noise on cost and revenue lines."""
    d = drivers("Base").copy()
    d.update(new_base=25, new_growth=0.015, churn=0.021, arpa=312, host_cost=15.2, cac=660, prof=4900, other=6300, tools_per_head=187)
    hires = [0, 0, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0]       # first hire slipped from June to July
    r = engine(d, hires)
    rng = np.random.default_rng(7)
    out = {}
    for key in ["sub_rev", "onb_rev", "hosting", "support", "payfees", "salaries", "marketing", "tools", "rent", "prof", "other"]:
        noise = 1 + rng.uniform(-0.015, 0.015, N_ACTUAL)
        out[key] = [round(float(v * k), 0) for v, k in zip(r[key][:N_ACTUAL], noise)]
    out["new"] = [int(v) for v in r["new"][:N_ACTUAL]]
    out["churned"] = [int(v) for v in r["churned"][:N_ACTUAL]]
    out["close_cust"] = [int(v) for v in r["close_cust"][:N_ACTUAL]]
    out["close_cash"] = [round(float(v), 0) for v in r["close_cash"][:N_ACTUAL]]
    return out


if __name__ == "__main__":
    for s in SCENARIOS:
        r = engine(drivers(s), HIRES[s])
        print(f"{s:9s} revenue {sum(r['revenue']):>12,.0f}  EBITDA {sum(r['ebitda']):>11,.0f}  closing customers {r['close_cust'][-1]:>5}  Mar cash {r['close_cash'][-1]:>11,.0f}  min cash {min(r['close_cash']):>11,.0f}")
    a = synthetic_actuals()
    b = engine(drivers('Base'), HIRES['Base'])
    print("actual vs budget revenue Apr-Sep:", round(sum(a['sub_rev']) + sum(a['onb_rev'])), round(sum(b['revenue'][:6])))
