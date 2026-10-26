"""Build Northlight_Scenario_Budget_Model.xlsx with live Excel formulas (xlsxwriter).

    python src/build_workbook.py

Sheets: Cover | Dashboard | Assumptions | Budget_vs_Actual | Actuals | Budget_Locked | Model_Base | Model_Upside | Model_Downside |
        Scenarios | ChartData | Checks
All calculations are Excel formulas. Python only writes them, and supplies the *locked* budget numbers and the fictional actuals.
"""
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

import model_spec as ms

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "Northlight_Scenario_Budget_Model.xlsx"

NAVY, BLUE, ORANGE, GREEN, RED, AMBER, GREY, LIGHT = "#1F3B5C", "#2F5D8A", "#E07B39", "#4C9F70", "#B5484C", "#C9A227", "#7A7A7A", "#F2F5F9"
MCOLS = [xl_col_to_name(2 + i) for i in range(12)]      # C..N
FY = "O"
SCEN_COL = {"Base": "C", "Upside": "D", "Downside": "E"}
GBP = '£#,##0;(£#,##0);"-"'

# ----------------------------------------------------------------------------------
# engine layout:  key, label, unit, number format, fy aggregation, style
#   agg: sum | first | last | none | custom (formula given in FY_CUSTOM)
# ----------------------------------------------------------------------------------
ENGINE = [
    ("h", "CUSTOMERS"),
    ("idx", "Month number", "#", "int", "none", "muted"),
    ("open_cust", "Opening customers", "customers", "int", "first", ""),
    ("new", "New customers", "customers", "int", "sum", ""),
    ("churned", "Churned customers", "customers", "int", "sum", ""),
    ("close_cust", "Closing customers", "customers", "int", "last", "bold"),
    ("avg_cust", "Average customers", "customers", "num1", "avg", ""),
    ("arpa", "ARPA (after price increase)", "£ / month", "gbp2", "avg", ""),
    ("h", "PROFIT AND LOSS"),
    ("sub_rev", "Subscription revenue", "£", "gbp", "sum", ""),
    ("onb_rev", "Onboarding revenue", "£", "gbp", "sum", ""),
    ("revenue", "Total revenue", "£", "gbp", "sum", "bold"),
    ("hosting", "Hosting", "£", "gbp", "sum", ""),
    ("support", "Customer support", "£", "gbp", "sum", ""),
    ("payfees", "Payment fees", "£", "gbp", "sum", ""),
    ("cogs", "Total cost of sales", "£", "gbp", "sum", "bold"),
    ("gross_profit", "Gross profit", "£", "gbp", "sum", "bold"),
    ("gm", "Gross margin %", "%", "pct1", "custom", "muted"),
    ("hc_open", "Opening headcount", "FTE", "int", "first", "muted"),
    ("hires", "New hires", "FTE", "int", "sum", "muted"),
    ("hc_close", "Closing headcount", "FTE", "int", "last", "muted"),
    ("salaries", "Salaries and on-costs", "£", "gbp", "sum", ""),
    ("marketing", "Marketing (CAC x new customers)", "£", "gbp", "sum", ""),
    ("tools", "Software and tools", "£", "gbp", "sum", ""),
    ("rent", "Rent and facilities", "£", "gbp", "sum", ""),
    ("prof", "Professional fees", "£", "gbp", "sum", ""),
    ("other", "Other G&A", "£", "gbp", "sum", ""),
    ("opex", "Total operating expenses", "£", "gbp", "sum", "bold"),
    ("ebitda", "EBITDA", "£", "gbp", "sum", "total"),
    ("ebitda_m", "EBITDA margin %", "%", "pct1", "custom", "muted"),
    ("da", "Depreciation", "£", "gbp", "sum", ""),
    ("ebit", "EBIT", "£", "gbp", "sum", "bold"),
    ("tax", "Corporation tax charge", "£", "gbp", "sum", ""),
    ("net_income", "Net income", "£", "gbp", "sum", "total"),
    ("h", "CASH FLOW"),
    ("receipts", "Cash received from customers", "£", "gbp", "sum", ""),
    ("cogs_paid", "Cost of sales paid", "£", "gbp", "sum", ""),
    ("opex_paid", "Operating expenses paid", "£", "gbp", "sum", ""),
    ("tax_paid", "Tax paid (month after charge)", "£", "gbp", "sum", ""),
    ("capex", "Capital expenditure", "£", "gbp", "sum", ""),
    ("net_cash", "Net cash flow", "£", "gbp", "sum", "bold"),
    ("open_cash", "Opening cash", "£", "gbp", "first", ""),
    ("close_cash", "Closing cash", "£", "gbp", "last", "total"),
    ("h", "KPIs"),
    ("mrr", "Monthly recurring revenue (closing)", "£", "gbp", "last", ""),
    ("arr", "Annualised recurring revenue (closing)", "£", "gbp", "last", ""),
]
ROW = {}
r = 5
for e in ENGINE:
    if e[0] != "h":
        ROW[e[0]] = r
    r += 1
LAST_ENGINE_ROW = r - 1
CHECK_START = LAST_ENGINE_ROW + 3

# P&L lines used in variance analysis: key, label, income|cost, is_total
BVA_LINES = [
    ("sub_rev", "Subscription revenue", "income", False), ("onb_rev", "Onboarding revenue", "income", False),
    ("revenue", "Total revenue", "income", True),
    ("hosting", "Hosting", "cost", False), ("support", "Customer support", "cost", False), ("payfees", "Payment fees", "cost", False),
    ("cogs", "Total cost of sales", "cost", True), ("gross_profit", "Gross profit", "income", True),
    ("salaries", "Salaries and on-costs", "cost", False), ("marketing", "Marketing", "cost", False), ("tools", "Software and tools", "cost", False),
    ("rent", "Rent and facilities", "cost", False), ("prof", "Professional fees", "cost", False), ("other", "Other G&A", "cost", False),
    ("opex", "Total operating expenses", "cost", True), ("ebitda", "EBITDA", "income", True),
]
ACT_INPUTS = ["new", "churned", "close_cust", "sub_rev", "onb_rev", "hosting", "support", "payfees", "salaries", "marketing", "tools", "rent",
              "prof", "other", "close_cash"]


class Book:
    def __init__(self):
        self.wb = xlsxwriter.Workbook(str(OUT))
        wb = self.wb
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        mk = lambda **k: wb.add_format({**base, **k})
        self.f = {
            "title": mk(bold=True, font_size=18, font_color=NAVY), "sub": mk(font_color=GREY, italic=True),
            "h": mk(bold=True, font_color="white", bg_color=NAVY, align="center", border=1, border_color=NAVY),
            "hl": mk(bold=True, font_color="white", bg_color=NAVY, border=1, border_color=NAVY),
            "sec": mk(bold=True, font_color=NAVY, bg_color="#DCE6F1", bottom=1), "text": mk(), "wrap": mk(text_wrap=True, valign="top"),
            "muted": mk(font_color=GREY), "bold": mk(bold=True),
            "note": mk(font_color=GREY, italic=True, text_wrap=True, valign="top"),
        }
        nf = {"gbp": GBP, "gbp2": '£#,##0.00', "int": "#,##0", "num1": "#,##0.0", "pct1": "0.0%", "pct0": "0%"}
        for kind, num in nf.items():
            self.f[kind] = mk(num_format=num)
            self.f[kind + "_b"] = mk(num_format=num, bold=True)
            self.f[kind + "_t"] = mk(num_format=num, bold=True, top=1, bottom=6, bg_color=LIGHT)
            self.f[kind + "_m"] = mk(num_format=num, font_color=GREY)
            self.f[kind + "_in"] = mk(num_format=num, font_color="#0000FF", bg_color="#FFF9E5", border=1, border_color="#D9D9D9")
            self.f[kind + "_lock"] = mk(num_format=num, font_color="#0000FF", bg_color="#EDEDED")
        self.f["label_t"] = mk(bold=True, top=1, bottom=6, bg_color=LIGHT)
        self.f["input_text"] = mk(font_color="#0000FF", bg_color="#FFF9E5", border=1, border_color="#D9D9D9", bold=True, align="center")

    def style(self, kind, style):
        suffix = {"bold": "_b", "total": "_t", "muted": "_m", "in": "_in", "lock": "_lock", "": ""}[style]
        return self.f[kind + suffix]


def build():
    B = Book()
    wb, f = B.wb, B.f
    sheets = {n: wb.add_worksheet(n) for n in ["Cover", "Dashboard", "Assumptions", "Budget_vs_Actual", "Actuals", "Budget_Locked",
                                                "Model_Base", "Model_Upside", "Model_Downside", "Scenarios", "ChartData", "Checks"]}
    for n, color in [("Cover", NAVY), ("Dashboard", ORANGE), ("Assumptions", "#0000FF"), ("Actuals", "#0000FF"), ("Budget_vs_Actual", GREEN),
                     ("Budget_Locked", GREY), ("Model_Base", BLUE), ("Model_Upside", BLUE), ("Model_Downside", BLUE), ("Scenarios", GREEN),
                     ("ChartData", GREY), ("Checks", RED)]:
        sheets[n].set_tab_color(color)
        sheets[n].hide_gridlines(2)

    # =========================================================== Assumptions
    ws = sheets["Assumptions"]
    ws.set_column("A:A", 46); ws.set_column("B:B", 20); ws.set_column("C:N", 11.5); ws.set_column("G:G", 11.5)
    ws.write("A1", "Assumptions", f["title"])
    ws.write("A2", "Blue cells on yellow are inputs. Every driver is documented in the rationale column. Upside / Downside link to Base where a driver is shared.", f["sub"])
    for c, t in zip("ABCDEF", ["Driver", "Unit", "Base", "Upside", "Downside", "Live (selected)"]):
        ws.write(f"{c}4", t, f["hl"] if c in "AB" else f["h"])
    ws.merge_range("G4:N4", "Rationale / source (all illustrative)", f["hl"])
    AROW = {}
    row = 5
    for a in ms.ASSUMPTIONS:
        if a[0] == "section":
            ws.merge_range(row - 1, 0, row - 1, 13, a[1], f["sec"])
            row += 1
            continue
        key, label, unit, vb, vu, vd, fmt, note = a
        AROW[key] = row
        ws.write(row - 1, 0, label, f["text"]); ws.write(row - 1, 1, unit, f["muted"])
        ws.write(row - 1, 2, vb, B.style(fmt, "in"))
        for col, v in ((3, vu), (4, vd)):
            if v is None:
                ws.write_formula(row - 1, col, f"=$C{row}", B.style(fmt, ""))
            else:
                ws.write(row - 1, col, v, B.style(fmt, "in"))
        ws.write_formula(row - 1, 5, f"=CHOOSE(Scn_Idx,C{row},D{row},E{row})", B.style(fmt, "bold"))
        ws.merge_range(row - 1, 6, row - 1, 13, note, f["note"])
        row += 1
    row += 1
    ws.merge_range(row - 1, 0, row - 1, 13, "Hiring plan: new hires starting in each month (FTE)", f["sec"])
    row += 1
    ws.write(row - 1, 0, "Scenario", f["hl"]); ws.write(row - 1, 1, "", f["hl"])
    for i, m in enumerate(ms.MONTHS):
        ws.write(row - 1, 2 + i, m, f["h"])
    HIRE_ROW = {}
    for s in ms.SCENARIOS:
        row += 1
        HIRE_ROW[s] = row
        ws.write(row - 1, 0, f"{s} case hires", f["text"]); ws.write(row - 1, 1, "FTE", f["muted"])
        for i, v in enumerate(ms.HIRES[s]):
            ws.write(row - 1, 2 + i, v, f["int_in"])
    row += 2
    ws.merge_range(row - 1, 0, row - 1, 13, "Variance reporting thresholds", f["sec"])
    row += 1
    AROW["amber"] = row
    ws.write(row - 1, 0, "Amber threshold (adverse variance as % of budget)", f["text"]); ws.write(row - 1, 1, "%", f["muted"]); ws.write(row - 1, 2, 0.02, f["pct1_in"])
    ws.merge_range(row - 1, 6, row - 1, 13, "Adverse variance beyond this is flagged 'Watch'.", f["note"])
    row += 1
    AROW["red"] = row
    ws.write(row - 1, 0, "Red threshold (adverse variance as % of budget)", f["text"]); ws.write(row - 1, 1, "%", f["muted"]); ws.write(row - 1, 2, 0.05, f["pct1_in"])
    ws.merge_range(row - 1, 6, row - 1, 13, "Adverse variance beyond this is flagged 'Off track'.", f["note"])
    row += 1
    ws.write(row, 0, "Scenario index (calculated from the Dashboard selector)", f["muted"])
    ws.write_formula(row, 2, '=MATCH(Scenario,{"Base","Upside","Downside"},0)', f["int"])
    wb.define_name("Scn_Idx", f"=Assumptions!$C${row + 1}")
    ws.freeze_panes(4, 2)

    def A(key, scen):
        return f"Assumptions!${SCEN_COL[scen]}${AROW[key]}"

    # =========================================================== engine sheets
    FY_CUSTOM = {
        "gm": lambda: f"=IFERROR({FY}{ROW['gross_profit']}/{FY}{ROW['revenue']},0)",
        "ebitda_m": lambda: f"=IFERROR({FY}{ROW['ebitda']}/{FY}{ROW['revenue']},0)",
    }

    def engine_formulas(scen):
        a = lambda k: A(k, scen)
        R = ROW
        def fm(key, c, p, i):
            idx = f"{c}{R['idx']}"
            table = {
                "idx": None,
                "open_cust": f"={a('open_customers')}" if i == 0 else f"={p}{R['close_cust']}",
                "new": f"=ROUND({a('new_base')}*(1+{a('new_growth')})^({idx}-1),0)",
                "churned": f"=ROUND({c}{R['open_cust']}*{a('churn')},0)",
                "close_cust": f"={c}{R['open_cust']}+{c}{R['new']}-{c}{R['churned']}",
                "avg_cust": f"=({c}{R['open_cust']}+{c}{R['close_cust']})/2",
                "arpa": f"={a('arpa')}*IF({idx}>={a('price_month')},1+{a('price_inc')},1)",
                "sub_rev": f"={c}{R['avg_cust']}*{c}{R['arpa']}",
                "onb_rev": f"={c}{R['new']}*{a('onboard_fee')}",
                "revenue": f"={c}{R['sub_rev']}+{c}{R['onb_rev']}",
                "hosting": f"={c}{R['avg_cust']}*{a('host_cost')}",
                "support": f"={c}{R['avg_cust']}*{a('support_cost')}",
                "payfees": f"={c}{R['revenue']}*{a('pay_fee')}",
                "cogs": f"={c}{R['hosting']}+{c}{R['support']}+{c}{R['payfees']}",
                "gross_profit": f"={c}{R['revenue']}-{c}{R['cogs']}",
                "gm": f"=IFERROR({c}{R['gross_profit']}/{c}{R['revenue']},0)",
                "hc_open": f"={a('hc_open')}" if i == 0 else f"={p}{R['hc_close']}",
                "hires": f"=Assumptions!{c}{HIRE_ROW[scen]}",
                "hc_close": f"={c}{R['hc_open']}+{c}{R['hires']}",
                "salaries": f"={c}{R['hc_close']}*{a('salary')}/12*(1+{a('oncost')})",
                "marketing": f"={c}{R['new']}*{a('cac')}",
                "tools": f"={c}{R['hc_close']}*{a('tools_per_head')}",
                "rent": f"={a('rent')}", "prof": f"={a('prof')}", "other": f"={a('other')}",
                "opex": f"={c}{R['salaries']}+{c}{R['marketing']}+{c}{R['tools']}+{c}{R['rent']}+{c}{R['prof']}+{c}{R['other']}",
                "ebitda": f"={c}{R['gross_profit']}-{c}{R['opex']}",
                "ebitda_m": f"=IFERROR({c}{R['ebitda']}/{c}{R['revenue']},0)",
                "da": f"={a('dep_existing')}+{a('capex_m')}*({idx}-1)/{a('dep_life')}",
                "ebit": f"={c}{R['ebitda']}-{c}{R['da']}",
                "tax": f"=MAX(0,{c}{R['ebit']})*{a('tax_rate')}",
                "net_income": f"={c}{R['ebit']}-{c}{R['tax']}",
                "receipts": (f"={c}{R['revenue']}*{a('coll_in_month')}+{a('open_recv')}" if i == 0
                             else f"={c}{R['revenue']}*{a('coll_in_month')}+{p}{R['revenue']}*(1-{a('coll_in_month')})"),
                "cogs_paid": f"={c}{R['cogs']}", "opex_paid": f"={c}{R['opex']}",
                "tax_paid": f"={a('open_tax')}" if i == 0 else f"={p}{R['tax']}",
                "capex": f"={a('capex_m')}",
                "net_cash": f"={c}{R['receipts']}-{c}{R['cogs_paid']}-{c}{R['opex_paid']}-{c}{R['tax_paid']}-{c}{R['capex']}",
                "open_cash": f"={a('open_cash')}" if i == 0 else f"={p}{R['close_cash']}",
                "close_cash": f"={c}{R['open_cash']}+{c}{R['net_cash']}",
                "mrr": f"={c}{R['close_cust']}*{c}{R['arpa']}",
                "arr": f"={c}{R['mrr']}*12",
            }
            return table[key]
        return fm

    def fy_formula(key, agg):
        rr = ROW[key]
        if agg == "sum":
            return f"=SUM(C{rr}:N{rr})"
        if agg == "first":
            return f"=C{rr}"
        if agg == "last":
            return f"=N{rr}"
        if agg == "avg":
            return f"=AVERAGE(C{rr}:N{rr})"
        if agg == "custom":
            return FY_CUSTOM[key]()
        return None

    def write_engine(ws, title, subtitle, scen=None, static=None):
        ws.set_column("A:A", 36); ws.set_column("B:B", 11); ws.set_column("C:O", 11.5)
        ws.write("A1", title, f["title"]); ws.write("A2", subtitle, f["sub"])
        ws.write("A4", "£ unless stated", f["hl"]); ws.write("B4", "Unit", f["h"])
        for i, m in enumerate(ms.MONTHS):
            ws.write(3, 2 + i, m, f["h"])
        ws.write("O4", "FY 2026/27", f["h"])
        fm = engine_formulas(scen) if scen else None
        rr = 5
        for e in ENGINE:
            if e[0] == "h":
                ws.merge_range(rr - 1, 0, rr - 1, 14, e[1], f["sec"])
                rr += 1
                continue
            key, label, unit, kind, agg, style = e
            lab = f["label_t"] if style == "total" else (f["bold"] if style == "bold" else (f["muted"] if style == "muted" else f["text"]))
            ws.write(rr - 1, 0, label, lab); ws.write(rr - 1, 1, unit, f["muted"])
            for i, c in enumerate(MCOLS):
                p = MCOLS[i - 1] if i else None
                if key == "idx":
                    ws.write(rr - 1, 2 + i, i + 1, f["int_m"])
                elif static is not None:
                    ws.write(rr - 1, 2 + i, static[key][i], B.style(kind, "lock") if style != "total" else f[kind + "_t"])
                else:
                    ws.write_formula(rr - 1, 2 + i, fm(key, c, p, i), B.style(kind, style))
            fyf = fy_formula(key, agg)
            if fyf:
                ws.write_formula(rr - 1, 14, fyf, B.style(kind, style if style in ("bold", "total", "muted") else ""))
            rr += 1
        ws.freeze_panes(4, 2)
        return rr

    def add_engine_checks(ws, sname, row0):
        R = ROW
        checks = [
            ("Cash: opening + net cash flow = closing (FY)", f"=ROUND({FY}{R['open_cash']}+{FY}{R['net_cash']}-{FY}{R['close_cash']},2)"),
            ("Customers: opening + new - churned = closing (FY)", f"={FY}{R['open_cust']}+{FY}{R['new']}-{FY}{R['churned']}-{FY}{R['close_cust']}"),
            ("P&L: revenue - cost of sales - opex = EBITDA (FY)", f"=ROUND({FY}{R['revenue']}-{FY}{R['cogs']}-{FY}{R['opex']}-{FY}{R['ebitda']},2)"),
            ("P&L: EBITDA - depreciation - tax = net income (FY)", f"=ROUND({FY}{R['ebitda']}-{FY}{R['da']}-{FY}{R['tax']}-{FY}{R['net_income']},2)"),
            ("No Excel errors anywhere on the sheet", f"=SUMPRODUCT(--ISERROR(C5:{FY}{LAST_ENGINE_ROW}))"),
        ]
        ws.merge_range(row0 - 1, 0, row0 - 1, 14, "INTEGRITY CHECKS (each must be 0)", f["sec"])
        out = []
        for j, (lab, fml) in enumerate(checks):
            ws.write(row0 + j, 0, lab, f["text"])
            ws.write_formula(row0 + j, 2, fml, f["num1"])
            out.append((f"{sname}: {lab}", f"{sname}!$C${row0 + j + 1}"))
        return out

    all_checks = []
    for s in ms.SCENARIOS:
        name = f"Model_{s}"
        write_engine(sheets[name], f"{s} case: monthly model", f"Driven by the '{s}' column of Assumptions. Same layout on all three scenario sheets and on Budget_Locked.", scen=s)
        all_checks += add_engine_checks(sheets[name], name, CHECK_START)

    # =========================================================== budget (locked values)
    base_run = ms.engine(ms.drivers("Base"), ms.HIRES["Base"])
    base_run["gm"] = [g / r_ if r_ else 0 for g, r_ in zip(base_run["gross_profit"], base_run["revenue"])]
    base_run["ebitda_m"] = [e / r_ if r_ else 0 for e, r_ in zip(base_run["ebitda"], base_run["revenue"])]
    base_run["idx"] = list(range(1, 13))
    write_engine(sheets["Budget_Locked"], "Approved budget (locked values)",
                 "Hard-coded copy of the Base case at budget approval (1 Apr 2026). It does NOT move when Assumptions change: variances are always measured against this.",
                 static=base_run)
    sheets["Budget_Locked"].write("Q4", "Lock note", f["hl"])
    sheets["Budget_Locked"].set_column("Q:Q", 60)
    sheets["Budget_Locked"].write("Q5", "Grey cells are values, not formulas. To re-baseline the budget, copy Model_Base and paste as values here.", f["note"])

    # =========================================================== actuals
    ws = sheets["Actuals"]
    ws.set_column("A:A", 36); ws.set_column("B:B", 11); ws.set_column("C:I", 12)
    ws.write("A1", "Actuals: April to September 2026", f["title"])
    ws.write("A2", "Illustrative, fictional figures. Replace the blue inputs each month; everything else updates. Totals are formulas.", f["sub"])
    ws.write("A4", "£ unless stated", f["hl"]); ws.write("B4", "Unit", f["h"])
    for i in range(ms.N_ACTUAL):
        ws.write(3, 2 + i, ms.MONTHS[i], f["h"])
    ws.write("I4", "YTD", f["h"])
    act = ms.synthetic_actuals()
    AR = {}
    arow = 5
    layout = [("h", "CUSTOMERS"), ("new", "New customers", "customers", "int"), ("churned", "Churned customers", "customers", "int"),
              ("close_cust", "Closing customers", "customers", "int"), ("h", "REVENUE"), ("sub_rev", "Subscription revenue", "£", "gbp"),
              ("onb_rev", "Onboarding revenue", "£", "gbp"), ("revenue", "Total revenue", "£", "gbp"),
              ("h", "COST OF SALES"), ("hosting", "Hosting", "£", "gbp"), ("support", "Customer support", "£", "gbp"),
              ("payfees", "Payment fees", "£", "gbp"), ("cogs", "Total cost of sales", "£", "gbp"), ("gross_profit", "Gross profit", "£", "gbp"),
              ("h", "OPERATING EXPENSES"), ("salaries", "Salaries and on-costs", "£", "gbp"), ("marketing", "Marketing", "£", "gbp"),
              ("tools", "Software and tools", "£", "gbp"), ("rent", "Rent and facilities", "£", "gbp"), ("prof", "Professional fees", "£", "gbp"),
              ("other", "Other G&A", "£", "gbp"), ("opex", "Total operating expenses", "£", "gbp"), ("ebitda", "EBITDA", "£", "gbp"),
              ("h", "CASH"), ("close_cash", "Closing cash (bank)", "£", "gbp")]
    for e in layout:
        if e[0] == "h":
            ws.merge_range(arow - 1, 0, arow - 1, 8, e[1], f["sec"]); arow += 1; continue
        AR[e[0]] = arow; arow += 1
    calc = {"revenue": lambda c: f"={c}{AR['sub_rev']}+{c}{AR['onb_rev']}",
            "cogs": lambda c: f"={c}{AR['hosting']}+{c}{AR['support']}+{c}{AR['payfees']}",
            "gross_profit": lambda c: f"={c}{AR['revenue']}-{c}{AR['cogs']}",
            "opex": lambda c: f"=SUM({c}{AR['salaries']}:{c}{AR['other']})",
            "ebitda": lambda c: f"={c}{AR['gross_profit']}-{c}{AR['opex']}"}
    for e in layout:
        if e[0] == "h":
            continue
        key, label, unit, kind = e
        rr = AR[key]
        total = key in calc
        ws.write(rr - 1, 0, label, f["label_t"] if key in ("ebitda",) else (f["bold"] if total else f["text"])); ws.write(rr - 1, 1, unit, f["muted"])
        for i in range(ms.N_ACTUAL):
            c = MCOLS[i]
            if total:
                ws.write_formula(rr - 1, 2 + i, calc[key](c), B.style(kind, "total" if key == "ebitda" else "bold"))
            else:
                ws.write(rr - 1, 2 + i, act[key][i], B.style(kind, "in"))
        if key in ("close_cust", "close_cash"):
            ws.write_formula(rr - 1, 8, f"=H{rr}", B.style(kind, "bold"))
        else:
            ws.write_formula(rr - 1, 8, f"=SUM(C{rr}:H{rr})", B.style(kind, "total" if key == "ebitda" else "bold"))
    ws.freeze_panes(4, 2)

    # =========================================================== budget vs actual
    ws = sheets["Budget_vs_Actual"]
    ws.set_column("A:A", 44); ws.set_column("B:L", 13.5)
    ws.write("A1", "Budget vs actual", f["title"])
    ws.write("A2", "Variances are shown as favourable (+) / adverse (-) to profit. Outlook = YTD actual + forecast for the remaining months from the scenario selected on the Dashboard.", f["sub"])
    ws.write("A3", "Scenario used for the outlook:", f["bold"]); ws.write_formula("C3", "=Scenario", f["bold"])
    hdr = ["Line", "YTD budget", "YTD actual", "YTD var £ (F/A)", "YTD var %", "Status", "FY budget", "Forecast Oct-Mar", "FY outlook", "Outlook var £ (F/A)", "Outlook var %"]
    for j, h in enumerate(hdr):
        ws.write(4, j, h, f["hl"] if j == 0 else f["h"])
    ws.set_row(4, 30)
    BVA = {}
    for i, (key, label, typ, tot) in enumerate(BVA_LINES):
        rr = 6 + i
        BVA[key] = rr
        rb = ROW[key]
        fmt_n = "gbp_b" if tot else "gbp"
        ws.write(rr - 1, 0, label, f["bold"] if tot else f["text"])
        ws.write_formula(rr - 1, 1, f"=SUM(Budget_Locked!C{rb}:H{rb})", f[fmt_n])
        ws.write_formula(rr - 1, 2, f"=Actuals!I{AR[key]}", f[fmt_n])
        sign = "C{r}-B{r}" if typ == "income" else "B{r}-C{r}"
        ws.write_formula(rr - 1, 3, "=" + sign.format(r=rr), f[fmt_n])
        ws.write_formula(rr - 1, 4, f"=IFERROR(D{rr}/ABS(B{rr}),0)", f["pct1_b" if tot else "pct1"])
        ws.write_formula(rr - 1, 5, f'=IF(E{rr}>=0,"Favourable",IF(E{rr}>=-Assumptions!$C${AROW["amber"]},"On track",IF(E{rr}>=-Assumptions!$C${AROW["red"]},"Watch","Off track")))', f["bold"])
        ws.write_formula(rr - 1, 6, f"=Budget_Locked!O{rb}", f[fmt_n])
        ws.write_formula(rr - 1, 7, f"=CHOOSE(Scn_Idx,SUM(Model_Base!I{rb}:N{rb}),SUM(Model_Upside!I{rb}:N{rb}),SUM(Model_Downside!I{rb}:N{rb}))", f[fmt_n])
        ws.write_formula(rr - 1, 8, f"=C{rr}+H{rr}", f[fmt_n])
        osign = "I{r}-G{r}" if typ == "income" else "G{r}-I{r}"
        ws.write_formula(rr - 1, 9, "=" + osign.format(r=rr), f[fmt_n])
        ws.write_formula(rr - 1, 10, f"=IFERROR(J{rr}/ABS(G{rr}),0)", f["pct1_b" if tot else "pct1"])
    last = 6 + len(BVA_LINES) - 1
    red_f = wb.add_format({"font_color": RED, "bold": True}); grn_f = wb.add_format({"font_color": "#2E7D32", "bold": True})
    for col in ("D", "E", "J", "K"):
        ws.conditional_format(f"{col}6:{col}{last}", {"type": "cell", "criteria": "<", "value": 0, "format": red_f})
        ws.conditional_format(f"{col}6:{col}{last}", {"type": "cell", "criteria": ">=", "value": 0, "format": grn_f})
    for txt, bg, fc in (("Off track", "#F4CCCC", "#9C0006"), ("Watch", "#FFE8B3", "#7F6000"), ("On track", "#E2EFDA", "#375623"), ("Favourable", "#C6E0B4", "#375623")):
        ws.conditional_format(f"F6:F{last}", {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": wb.add_format({"bg_color": bg, "font_color": fc, "bold": True})})
    ws.write(last + 1, 0, "Key operating metrics at 30 September", f["sec"])
    for j in range(1, 11):
        ws.write(last + 1, j, "", f["sec"])
    for j, h in enumerate(["Metric", "Budget", "Actual", "Var (F/A)", "Var %"]):
        ws.write(last + 2, j, h, f["hl"] if j == 0 else f["h"])
    km = [("Closing customers", "close_cust", "close_cust", "int", +1), ("Closing cash (£)", "close_cash", "close_cash", "gbp", +1)]
    KM = {}
    for j, (lab, bk, ak, kind, sg) in enumerate(km):
        rr = last + 4 + j
        KM[bk] = rr
        ws.write(rr - 1, 0, lab, f["text"])
        ws.write_formula(rr - 1, 1, f"=Budget_Locked!H{ROW[bk]}", f[kind])
        ws.write_formula(rr - 1, 2, f"=Actuals!H{AR[ak]}", f[kind])
        ws.write_formula(rr - 1, 3, f"=C{rr}-B{rr}", f[kind])
        ws.write_formula(rr - 1, 4, f"=IFERROR(D{rr}/ABS(B{rr}),0)", f["pct1"])
    rr = last + 6
    ws.write(rr - 1, 0, "Total new customers Apr-Sep", f["text"])
    ws.write_formula(rr - 1, 1, f"=SUM(Budget_Locked!C{ROW['new']}:H{ROW['new']})", f["int"]); ws.write_formula(rr - 1, 2, f"=Actuals!I{AR['new']}", f["int"])
    ws.write_formula(rr - 1, 3, f"=C{rr}-B{rr}", f["int"]); ws.write_formula(rr - 1, 4, f"=IFERROR(D{rr}/ABS(B{rr}),0)", f["pct1"])
    KM["new"] = rr
    rr += 1
    ws.write(rr - 1, 0, "Total churned customers Apr-Sep (lower is better)", f["text"])
    ws.write_formula(rr - 1, 1, f"=SUM(Budget_Locked!C{ROW['churned']}:H{ROW['churned']})", f["int"]); ws.write_formula(rr - 1, 2, f"=Actuals!I{AR['churned']}", f["int"])
    ws.write_formula(rr - 1, 3, f"=B{rr}-C{rr}", f["int"]); ws.write_formula(rr - 1, 4, f"=IFERROR(D{rr}/ABS(B{rr}),0)", f["pct1"])
    KM["churned"] = rr
    ws.conditional_format(f"D{last+4}:E{rr}", {"type": "cell", "criteria": "<", "value": 0, "format": red_f})
    ws.conditional_format(f"D{last+4}:E{rr}", {"type": "cell", "criteria": ">=", "value": 0, "format": grn_f})
    ws.freeze_panes(5, 1)

    # =========================================================== scenarios
    ws = sheets["Scenarios"]
    ws.set_column("A:A", 34); ws.set_column("B:H", 14)
    ws.write("A1", "Scenario comparison: FY 2026/27", f["title"])
    ws.write("A2", "Full-year Base / Upside / Downside side by side, against the locked budget. Scenarios apply to all 12 months; actuals for Apr-Sep are on Budget_vs_Actual.", f["sub"])
    for j, h in enumerate(["Metric", "Budget (locked)", "Base", "Upside", "Downside", "Upside vs Base", "Downside vs Base"]):
        ws.write(3, j, h, f["hl"] if j == 0 else f["h"])
    rows_s = [("Total revenue", "revenue", "gbp"), ("Gross profit", "gross_profit", "gbp"), ("Gross margin %", "gm", "pct1"), ("Total operating expenses", "opex", "gbp"),
              ("EBITDA", "ebitda", "gbp"), ("EBITDA margin %", "ebitda_m", "pct1"), ("Net income", "net_income", "gbp"), ("Closing customers (Mar)", "close_cust", "int"),
              ("Closing ARR (Mar)", "arr", "gbp"), ("Closing headcount (Mar)", "hc_close", "int"), ("Closing cash (Mar)", "close_cash", "gbp")]
    SC = {}
    for i, (lab, key, kind) in enumerate(rows_s):
        rr = 5 + i
        SC[key] = rr
        ws.write(rr - 1, 0, lab, f["bold"] if key in ("ebitda", "close_cash") else f["text"])
        ws.write_formula(rr - 1, 1, f"=Budget_Locked!O{ROW[key]}", f[kind])
        for j, s in enumerate(ms.SCENARIOS):
            ws.write_formula(rr - 1, 2 + j, f"=Model_{s}!O{ROW[key]}", f[kind])
        diff_kind = kind if kind != "pct1" else "pct1"
        ws.write_formula(rr - 1, 5, f"=D{rr}-C{rr}", f[diff_kind]); ws.write_formula(rr - 1, 6, f"=E{rr}-C{rr}", f[diff_kind])
    rr = 5 + len(rows_s) + 1
    ws.write(rr - 1, 0, "Lowest month-end cash", f["bold"])
    ws.write_formula(rr - 1, 1, f"=MIN(Budget_Locked!C{ROW['close_cash']}:N{ROW['close_cash']})", f["gbp_b"])
    for j, s in enumerate(ms.SCENARIOS):
        ws.write_formula(rr - 1, 2 + j, f"=MIN(Model_{s}!C{ROW['close_cash']}:N{ROW['close_cash']})", f["gbp_b"])
    SC["min_cash"] = rr
    rr += 1
    ws.write(rr - 1, 0, "Month of lowest cash", f["text"])
    ws.write_formula(rr - 1, 1, f"=INDEX(Budget_Locked!$C$4:$N$4,MATCH(B{rr-1},Budget_Locked!C{ROW['close_cash']}:N{ROW['close_cash']},0))", f["text"])
    for j, s in enumerate(ms.SCENARIOS):
        col = xl_col_to_name(2 + j)
        ws.write_formula(rr - 1, 2 + j, f"=INDEX(Model_{s}!$C$4:$N$4,MATCH({col}{rr-1},Model_{s}!C{ROW['close_cash']}:N{ROW['close_cash']},0))", f["text"])
    rr += 1
    ws.write(rr - 1, 0, "Cash runway at Mar burn (months)", f["text"])
    for j, s in enumerate(ms.SCENARIOS):
        col = xl_col_to_name(2 + j)
        ws.write_formula(rr - 1, 2 + j, f'=IF(Model_{s}!N{ROW["net_cash"]}>=0,"Cash positive",Model_{s}!N{ROW["close_cash"]}/-Model_{s}!N{ROW["net_cash"]})', f["num1"])
    SC["runway"] = rr
    ws.merge_range(rr + 1, 0, rr + 1, 6, "Read-across: Upside and Downside differ from Base in new-customer volume and growth, churn, ARPA, price increase, hosting cost, CAC and hiring (see Assumptions).", f["note"])
    ws.set_row(rr + 1, 30)

    # =========================================================== chart data
    ws = sheets["ChartData"]
    ws.set_column("A:A", 12); ws.set_column("B:L", 15)
    ws.write("A1", "Chart data (feeds the Dashboard; all formulas)", f["title"])
    heads = ["Month", "Budget revenue", "Actual revenue", "Forecast revenue (selected)", "EBITDA Base", "EBITDA Upside", "EBITDA Downside", "Budget EBITDA",
             "Cash Base", "Cash Upside", "Cash Downside", "Budget cash"]
    for j, h in enumerate(heads):
        ws.write(3, j, h, f["h"])
    ws.set_row(3, 30)
    for i, c in enumerate(MCOLS):
        rr = 5 + i
        ws.write_formula(rr - 1, 0, f"=Model_Base!{c}4", f["text"])
        ws.write_formula(rr - 1, 1, f"=Budget_Locked!{c}{ROW['revenue']}", f["gbp"])
        ws.write_formula(rr - 1, 2, f"=Actuals!{c}{AR['revenue']}" if i < ms.N_ACTUAL else "=NA()", f["gbp"])
        ws.write_formula(rr - 1, 3, f"=CHOOSE(Scn_Idx,Model_Base!{c}{ROW['revenue']},Model_Upside!{c}{ROW['revenue']},Model_Downside!{c}{ROW['revenue']})" if i >= ms.N_ACTUAL else "=NA()", f["gbp"])
        for j, s in enumerate(ms.SCENARIOS):
            ws.write_formula(rr - 1, 4 + j, f"=Model_{s}!{c}{ROW['ebitda']}", f["gbp"])
            ws.write_formula(rr - 1, 8 + j, f"=Model_{s}!{c}{ROW['close_cash']}", f["gbp"])
        ws.write_formula(rr - 1, 7, f"=Budget_Locked!{c}{ROW['ebitda']}", f["gbp"])
        ws.write_formula(rr - 1, 11, f"=Budget_Locked!{c}{ROW['close_cash']}", f["gbp"])
    ws.write("A19", "Note: #N/A is intentional (charts skip those points).", f["note"])
    ws.write("A21", "YTD variance by line (F/A £)", f["bold"])
    chart_lines = [k for k in BVA if k in ("sub_rev", "onb_rev", "hosting", "support", "payfees", "salaries", "marketing", "tools", "rent", "prof", "other")]
    for j, k in enumerate(chart_lines):
        ws.write_formula(21 + j, 0, f"=Budget_vs_Actual!A{BVA[k]}", f["text"])
        ws.write_formula(21 + j, 1, f"=Budget_vs_Actual!D{BVA[k]}", f["gbp"])
    cl_first, cl_last = 22, 21 + len(chart_lines)

    # =========================================================== checks
    ws = sheets["Checks"]
    ws.set_column("A:A", 70); ws.set_column("B:C", 14); ws.set_column("D:D", 12)
    ws.write("A1", "Model integrity checks", f["title"])
    ws.write("A2", "Every check must be 0 (within tolerance). The overall status is shown on the Cover and Dashboard.", f["sub"])
    for j, h in enumerate(["Check", "Value", "Tolerance", "Status"]):
        ws.write(3, j, h, f["hl"] if j == 0 else f["h"])
    checks = list(all_checks)
    row0 = 5
    for j, (lab, ref) in enumerate(checks):
        rr = row0 + j
        ws.write(rr - 1, 0, lab, f["text"])
        ws.write_formula(rr - 1, 1, f"={ref}", f["num1"])
        ws.write(rr - 1, 2, 0.01, f["num1"])
        ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    rr = row0 + len(checks)
    # roll-forward of actual customers month by month
    ws.write(rr - 1, 0, "Actuals: customer roll-forward (Apr opening + new - churned = closing, then month on month)", f["text"])
    roll = (f"=ABS(Assumptions!$C${AROW['open_customers']}+Actuals!C{AR['new']}-Actuals!C{AR['churned']}-Actuals!C{AR['close_cust']})"
            + "".join(f"+ABS(Actuals!{MCOLS[k-1]}{AR['close_cust']}+Actuals!{MCOLS[k]}{AR['new']}-Actuals!{MCOLS[k]}{AR['churned']}-Actuals!{MCOLS[k]}{AR['close_cust']})" for k in range(1, ms.N_ACTUAL)))
    ws.write_formula(rr - 1, 1, roll, f["num1"]); ws.write(rr - 1, 2, 0.01, f["num1"]); ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    rr += 1
    ws.write(rr - 1, 0, "Budget_vs_Actual: YTD budget revenue ties to the sum of Budget_Locked months", f["text"])
    ws.write_formula(rr - 1, 1, f"=Budget_vs_Actual!B{BVA['revenue']}-SUM(Budget_Locked!C{ROW['revenue']}:H{ROW['revenue']})", f["num1"]); ws.write(rr - 1, 2, 0.01, f["num1"])
    ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    rr += 1
    ws.write(rr - 1, 0, "Budget_vs_Actual: actual YTD EBITDA = revenue - cost of sales - opex", f["text"])
    ws.write_formula(rr - 1, 1, f"=Budget_vs_Actual!C{BVA['ebitda']}-(Budget_vs_Actual!C{BVA['revenue']}-Budget_vs_Actual!C{BVA['cogs']}-Budget_vs_Actual!C{BVA['opex']})", f["num1"]); ws.write(rr - 1, 2, 0.01, f["num1"])
    ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    rr += 1
    ws.write(rr - 1, 0, "Scenario selector is valid (Base, Upside or Downside)", f["text"])
    ws.write_formula(rr - 1, 1, '=IF(ISNUMBER(Scn_Idx),0,1)', f["num1"]); ws.write(rr - 1, 2, 0, f["num1"]); ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    rr += 1
    ws.write(rr - 1, 0, "Dashboard / Budget_vs_Actual / Scenarios contain no Excel errors", f["text"])
    ws.write_formula(rr - 1, 1, "=SUMPRODUCT(--ISERROR(Budget_vs_Actual!A1:K40))+SUMPRODUCT(--ISERROR(Scenarios!A1:G25))+SUMPRODUCT(--ISERROR(Dashboard!A5:Q60))", f["num1"])
    ws.write(rr - 1, 2, 0, f["num1"]); ws.write_formula(rr - 1, 3, f'=IF(ABS(B{rr})<=C{rr},"PASS","FAIL")', f["bold"])
    last_check = rr
    rr += 2
    ws.write(rr - 1, 0, "OVERALL STATUS", f["bold"])
    ws.write_formula(rr - 1, 1, f'=IF(COUNTIF(D{row0}:D{last_check},"FAIL")=0,"All checks pass",COUNTIF(D{row0}:D{last_check},"FAIL")&" check(s) failing")', f["bold"])
    wb.define_name("Check_Status", f"=Checks!$B${rr}")
    rr += 2
    ws.write(rr - 1, 0, "Information (not an error): has the Base case moved since the budget was locked?", f["text"])
    ws.write_formula(rr - 1, 1, f'=IF(SUMPRODUCT(ABS(ROUND(Model_Base!C{ROW["revenue"]}:N{ROW["revenue"]}-Budget_Locked!C{ROW["revenue"]}:N{ROW["revenue"]},0)))+SUMPRODUCT(ABS(ROUND(Model_Base!C{ROW["ebitda"]}:N{ROW["ebitda"]}-Budget_Locked!C{ROW["ebitda"]}:N{ROW["ebitda"]},0)))=0,"Base = locked budget","Base differs from locked budget")', f["bold"])
    ws.conditional_format(f"D{row0}:D{last_check}", {"type": "cell", "criteria": "==", "value": '"PASS"', "format": wb.add_format({"bg_color": "#E2EFDA", "font_color": "#375623", "bold": True})})
    ws.conditional_format(f"D{row0}:D{last_check}", {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": wb.add_format({"bg_color": "#F4CCCC", "font_color": "#9C0006", "bold": True})})

    # =========================================================== dashboard
    ws = sheets["Dashboard"]
    ws.set_column("A:A", 2); ws.set_column("B:Q", 11)
    ws.set_row(0, 8)
    ws.merge_range("B2:L2", "Northlight Analytics: FY 2026/27 budget, scenarios and variance", f["title"])
    ws.merge_range("B3:L3", "Fictional company, illustrative numbers. Actuals Apr-Sep 2026; forecast Oct-Mar from the selected scenario.", f["sub"])
    ws.write("N2", "Scenario", f["bold"])
    ws.merge_range("O2:P2", "Base", f["input_text"])
    ws.data_validation("O2", {"validate": "list", "source": ["Base", "Upside", "Downside"], "input_title": "Scenario", "input_message": "Choose Base, Upside or Downside"})
    wb.define_name("Scenario", "=Dashboard!$O$2")
    ws.write("N3", "Checks", f["bold"])
    ws.merge_range("O3:Q3", "", f["bold"]); ws.write_formula("O3", "=Check_Status", f["bold"])
    ws.conditional_format("O3", {"type": "text", "criteria": "begins with", "value": "All", "format": wb.add_format({"font_color": "#2E7D32", "bold": True})})
    ws.conditional_format("O3", {"type": "text", "criteria": "not containing", "value": "All", "format": wb.add_format({"font_color": RED, "bold": True})})

    tile_lab = wb.add_format({"font_name": "Calibri", "font_size": 9, "font_color": GREY, "bg_color": LIGHT, "align": "center", "top": 2, "top_color": BLUE})
    tile_val = wb.add_format({"font_name": "Calibri", "font_size": 20, "bold": True, "font_color": NAVY, "bg_color": LIGHT, "align": "center", "num_format": '£#,##0,"k";(£#,##0,"k")'})
    tile_val_i = wb.add_format({"font_name": "Calibri", "font_size": 20, "bold": True, "font_color": NAVY, "bg_color": LIGHT, "align": "center", "num_format": "#,##0"})
    tile_sub = wb.add_format({"font_name": "Calibri", "font_size": 9, "bg_color": LIGHT, "align": "center", "num_format": '+0.0%;-0.0%;0.0%'})
    tile_sub_g = wb.add_format({"font_name": "Calibri", "font_size": 9, "bg_color": LIGHT, "align": "center", "num_format": '+£#,##0,"k" "vs budget";-£#,##0,"k" "vs budget"'})
    tiles = [
        ("YTD revenue (Apr-Sep)", f"=Budget_vs_Actual!C{BVA['revenue']}", tile_val, f"=Budget_vs_Actual!E{BVA['revenue']}", tile_sub, "pct"),
        ("YTD EBITDA", f"=Budget_vs_Actual!C{BVA['ebitda']}", tile_val, f"=Budget_vs_Actual!D{BVA['ebitda']}", tile_sub_g, "gbp"),
        ("FY revenue outlook", f"=Budget_vs_Actual!I{BVA['revenue']}", tile_val, f"=Budget_vs_Actual!K{BVA['revenue']}", tile_sub, "pct"),
        ("FY EBITDA outlook", f"=Budget_vs_Actual!I{BVA['ebitda']}", tile_val, f"=Budget_vs_Actual!J{BVA['ebitda']}", tile_sub_g, "gbp"),
        ("Cash at 31 Mar (outlook)", f'=CHOOSE(Scn_Idx,Model_Base!N{ROW["close_cash"]},Model_Upside!N{ROW["close_cash"]},Model_Downside!N{ROW["close_cash"]})', tile_val,
         f"=Scenarios!B{SC['close_cash']}", tile_sub_g, "cash"),
        ("Customers at 30 Sep", f"=Budget_vs_Actual!C{KM['close_cust']}", tile_val_i, f"=Budget_vs_Actual!E{KM['close_cust']}", tile_sub, "pct"),
    ]
    for i, (lab, valf, vf, subf, sf, kind) in enumerate(tiles):
        c0 = 1 + i * 2 + (i // 3)  # columns B-C, D-E, F-G | H-I, J-K, L-M
        c0 = [1, 3, 5, 8, 10, 12][i]
        ws.merge_range(4, c0, 4, c0 + 1, lab, tile_lab)
        ws.merge_range(5, c0, 6, c0 + 1, "", vf); ws.write_formula(5, c0, valf, vf)
        ws.merge_range(7, c0, 7, c0 + 1, "", sf)
        if kind == "cash":
            ws.write_formula(7, c0, f'=INDEX(Model_Base!C4:N4,MATCH(MIN(CHOOSE(Scn_Idx,Model_Base!C{ROW["close_cash"]}:N{ROW["close_cash"]},Model_Upside!C{ROW["close_cash"]}:N{ROW["close_cash"]},Model_Downside!C{ROW["close_cash"]}:N{ROW["close_cash"]})),CHOOSE(Scn_Idx,Model_Base!C{ROW["close_cash"]}:N{ROW["close_cash"]},Model_Upside!C{ROW["close_cash"]}:N{ROW["close_cash"]},Model_Downside!C{ROW["close_cash"]}:N{ROW["close_cash"]}),0))&" is the low point"', tile_sub)
        else:
            ws.write_formula(7, c0, subf, sf)
    for cell_range in ("B8:C8", "D8:E8", "F8:G8", "I8:J8", "K8:L8", "M8:N8"):
        pass
    ws.set_row(5, 18); ws.set_row(6, 18)
    # comparison table
    ws.merge_range("B10:K10", "Scenario comparison: FY 2026/27", f["sec"])
    for j, h in enumerate(["", "", "", "Budget", "Base", "Upside", "Downside"]):
        pass
    heads = {1: "Metric", 6: "Budget", 7: "Base", 8: "Upside", 9: "Downside"}
    ws.merge_range("B11:F11", "Metric", f["hl"])
    for col, h in zip(["G", "H", "I", "J"], ["Budget", "Base", "Upside", "Downside"]):
        ws.write(f"{col}11", h, f["h"])
    comp = [("Total revenue", "revenue", "gbp"), ("EBITDA", "ebitda", "gbp"), ("EBITDA margin %", "ebitda_m", "pct1"), ("Closing customers (Mar)", "close_cust", "int"),
            ("Closing cash (Mar)", "close_cash", "gbp"), ("Lowest month-end cash", "min_cash", "gbp")]
    for i, (lab, key, kind) in enumerate(comp):
        rr = 12 + i
        ws.merge_range(f"B{rr}:F{rr}", lab, f["text"])
        for col, src in zip(["G", "H", "I", "J"], ["B", "C", "D", "E"]):
            ws.write_formula(f"{col}{rr}", f"=Scenarios!{src}{SC[key]}", f[kind])
    ws.merge_range("B19:F19", "Runway at March burn rate (months)", f["text"])
    # charts
    pal = [BLUE, ORANGE, GREEN]
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Budget", "categories": "=ChartData!$A$5:$A$16", "values": "=ChartData!$B$5:$B$16", "fill": {"color": "#B8C6D8"}, "gap": 60})
    ch.add_series({"name": "Actual", "categories": "=ChartData!$A$5:$A$16", "values": "=ChartData!$C$5:$C$16", "fill": {"color": BLUE}})
    ln = wb.add_chart({"type": "line"})
    ln.add_series({"name": "Forecast (selected scenario)", "categories": "=ChartData!$A$5:$A$16", "values": "=ChartData!$D$5:$D$16",
                   "line": {"color": ORANGE, "width": 2.5, "dash_type": "dash"}, "marker": {"type": "circle", "size": 5, "fill": {"color": ORANGE}, "border": {"color": ORANGE}}})
    ch.combine(ln)
    ch.set_title({"name": "Monthly revenue: budget vs actual and forecast", "name_font": {"size": 11, "color": NAVY}})
    ch.set_y_axis({"num_format": '£#,##0,"k"', "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}})
    ch.set_x_axis({"num_font": {"size": 8}}); ch.set_legend({"position": "bottom"}); ch.set_size({"width": 560, "height": 290})
    ws.insert_chart("B21", ch)

    ch2 = wb.add_chart({"type": "line"})
    for j, (s, colr) in enumerate(zip(ms.SCENARIOS, pal)):
        ch2.add_series({"name": s, "categories": "=ChartData!$A$5:$A$16", "values": f"=ChartData!${xl_col_to_name(4+j)}$5:${xl_col_to_name(4+j)}$16", "line": {"color": colr, "width": 2.25}})
    ch2.add_series({"name": "Budget", "categories": "=ChartData!$A$5:$A$16", "values": "=ChartData!$H$5:$H$16", "line": {"color": GREY, "width": 1.5, "dash_type": "dash"}})
    ch2.set_title({"name": "Monthly EBITDA by scenario", "name_font": {"size": 11, "color": NAVY}})
    ch2.set_y_axis({"num_format": '£#,##0,"k"', "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}})
    ch2.set_x_axis({"num_font": {"size": 8}, "label_position": "low"}); ch2.set_legend({"position": "bottom"}); ch2.set_size({"width": 560, "height": 290})
    ws.insert_chart("J21", ch2)

    ch3 = wb.add_chart({"type": "line"})
    for j, (s, colr) in enumerate(zip(ms.SCENARIOS, pal)):
        ch3.add_series({"name": s, "categories": "=ChartData!$A$5:$A$16", "values": f"=ChartData!${xl_col_to_name(8+j)}$5:${xl_col_to_name(8+j)}$16", "line": {"color": colr, "width": 2.25}})
    ch3.add_series({"name": "Budget", "categories": "=ChartData!$A$5:$A$16", "values": "=ChartData!$L$5:$L$16", "line": {"color": GREY, "width": 1.5, "dash_type": "dash"}})
    ch3.set_title({"name": "Month-end cash by scenario", "name_font": {"size": 11, "color": NAVY}})
    ch3.set_y_axis({"num_format": '£#,##0,"k"', "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}})
    ch3.set_x_axis({"num_font": {"size": 8}}); ch3.set_legend({"position": "bottom"}); ch3.set_size({"width": 560, "height": 290})
    ws.insert_chart("B36", ch3)

    ch4 = wb.add_chart({"type": "bar"})
    ch4.add_series({"name": "YTD variance (F/A)", "categories": f"=ChartData!$A${cl_first}:$A${cl_last}", "values": f"=ChartData!$B${cl_first}:$B${cl_last}",
                    "fill": {"color": BLUE}, "invert_if_negative": True, "invert_if_negative_color": RED,
                    "data_labels": {"value": True, "num_format": '£#,##0,"k";-£#,##0,"k"', "font": {"size": 8}}})
    ch4.set_title({"name": "YTD variance by line: favourable (+) / adverse (-)", "name_font": {"size": 11, "color": NAVY}})
    ch4.set_x_axis({"num_format": '£#,##0,"k"', "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}})
    ch4.set_y_axis({"reverse": True, "label_position": "low", "num_font": {"size": 8}}); ch4.set_legend({"none": True}); ch4.set_size({"width": 560, "height": 290})
    ws.insert_chart("J36", ch4)
    ws.merge_range("B52:Q52", "Reading the dashboard: green = favourable to profit, red = adverse. Change the scenario in O2 to re-forecast Oct-Mar; actuals and the locked budget never change.", f["note"])
    ws.set_landscape(); ws.fit_to_pages(1, 1); ws.set_paper(9)

    # runway row on dashboard (scenario sheet row)
    for col, src in zip(["H", "I", "J"], ["C", "D", "E"]):
        ws.write_formula(f"{col}19", f"=Scenarios!{src}{SC['runway']}", f["num1"])
    ws.write("G19", "n/a", f["muted"])

    # =========================================================== cover
    ws = sheets["Cover"]
    ws.set_column("A:A", 3); ws.set_column("B:B", 26); ws.set_column("C:C", 100)
    ws.write("B2", "Northlight Analytics Ltd: scenario and budget-vs-actual model", f["title"])
    ws.write("B3", "FY 2026/27 (1 Apr 2026 - 31 Mar 2027). A fictional UK B2B analytics-software company; every number is illustrative.", f["sub"])
    ws.write("B5", "Model status", f["bold"]); ws.write_formula("C5", "=Check_Status", f["bold"])
    ws.write("B6", "Scenario shown", f["bold"]); ws.write_formula("C6", "=Scenario", f["bold"])
    blocks = [
        ("Purpose", "Plan next year's P&L and cash under three scenarios, track actual performance against the approved budget, and show whether the full-year outlook still meets plan."),
        ("How to use", "1) Pick a scenario on the Dashboard (cell O2).  2) Review or change drivers on Assumptions (blue cells only).  3) Paste each month's actuals into Actuals.  4) Read Budget_vs_Actual and the Dashboard.  5) Confirm the Checks sheet says 'All checks pass'."),
        ("Sheets", "Dashboard: KPIs, charts, scenario table.\nAssumptions: every driver with its rationale; hiring plan; variance thresholds.\nBudget_vs_Actual: YTD and full-year outlook variances with status flags.\nActuals: monthly actual inputs (Apr-Sep).\nBudget_Locked: approved budget, hard-coded values.\nModel_Base / Model_Upside / Model_Downside: the monthly engine (identical layout) for each scenario.\nScenarios: full-year side-by-side comparison.\nChartData: helper ranges for charts.\nChecks: integrity tests."),
        ("Colour code", "Blue text on yellow = input you may change.  Black = formula (do not overwrite).  Blue on grey = locked value.  Green/red numbers = favourable/adverse variance."),
        ("Key logic", "Customers roll forward monthly (opening + new - churned). Subscription revenue = average customers x ARPA; a price increase applies from the chosen month. Costs scale with customers (hosting, support), new customers (marketing) and headcount (salaries, tools). 80% of revenue is collected in the month billed, 20% the next month. Tax is 25% of positive monthly EBIT, paid the month after. Capex is depreciated straight-line over 36 months."),
        ("Variance convention", "Variances are favourable (+) or adverse (-) to profit: for income lines actual minus budget; for cost lines budget minus actual. Status flags use the thresholds on Assumptions (amber 2%, red 5% of budget)."),
        ("Simplifications", "No VAT, no balance-sheet statement, no loss carry-forward, no FX or interest, annual rather than monthly price cohorts, and hires start at the beginning of the month. Churn is applied to opening customers only. These are deliberate to keep the model transparent."),
        ("Data", "Fictional. The 'actuals' for Apr-Sep were generated from a plausible alternative set of drivers (slower sales, higher churn and hosting cost, a delayed hire, a legal-fee overrun) plus +/-1.5% noise. Do not read them as real company results."),
    ]
    rr = 8
    for h, t in blocks:
        ws.write(rr - 1, 1, h, wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "valign": "top"})); ws.write(rr - 1, 2, t, f["wrap"])
        ws.set_row(rr - 1, 14 * sum(1 + len(line) // 120 for line in t.split(chr(10))) + 6)
        rr += 1
    ws.set_landscape(); ws.fit_to_pages(1, 1)

    for n in ("Budget_vs_Actual", "Scenarios", "Assumptions", "Actuals", "Budget_Locked", "Model_Base", "Model_Upside", "Model_Downside", "Checks"):
        sheets[n].set_landscape(); sheets[n].fit_to_pages(1, 0); sheets[n].set_paper(9)
    sheets["Cover"].activate()
    wb.close()
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
