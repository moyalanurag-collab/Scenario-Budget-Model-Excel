"""Independent check of the Excel formulas: compare what Excel calculated against the Python mirror of the model.

    powershell -File src/recalc_and_dump.ps1 -Path Northlight_Scenario_Budget_Model.xlsx -Dump dump.json
    python src/verify.py dump.json
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_workbook as bw  # noqa: E402
import model_spec as ms  # noqa: E402


def cell(dump, sheet, ref_row, col_letter):
    d = dump[sheet]
    c = ord(col_letter) - 64 if len(col_letter) == 1 else None
    return d["values"][ref_row - d["first_row"]][c - d["first_col"]]


ERR_CODES = ("#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#ERR")


def main(path: str, scenario: str = "Base") -> int:
    dump = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    fails = 0
    checked = 0
    keys = [k for k in bw.ROW if k not in ("idx", "gm", "ebitda_m")]
    for scen in ms.SCENARIOS:
        py = ms.engine(ms.drivers(scen), ms.HIRES[scen])
        for key in keys:
            for i, col in enumerate(bw.MCOLS):
                x = cell(dump, f"Model_{scen}", bw.ROW[key], col)
                checked += 1
                if abs(float(x) - py[key][i]) > 0.01:
                    fails += 1
                    if fails < 10:
                        print(f"MISMATCH {scen} {key} {col}: excel {x} vs python {py[key][i]}")
    print(f"engine cells compared: {checked}, mismatches: {fails}")

    # Budget = Base at lock, so the locked sheet must equal Model_Base while Assumptions are untouched
    drift = 0
    for key in ("revenue", "ebitda", "close_cash", "close_cust"):
        for col in bw.MCOLS:
            if abs(float(cell(dump, "Budget_Locked", bw.ROW[key], col)) - float(cell(dump, "Model_Base", bw.ROW[key], col))) > 0.01:
                drift += 1
    print(f"locked budget vs live Base differences: {drift}")

    # budget-vs-actual table (YTD and full-year outlook) for the selected scenario
    act = ms.synthetic_actuals()
    a = {k: [float(x) for x in v] for k, v in act.items()}
    n = ms.N_ACTUAL
    sums = {k: sum(a[k]) for k in ("sub_rev", "onb_rev", "hosting", "support", "payfees", "salaries", "marketing", "tools", "rent", "prof", "other")}
    sums["revenue"] = sums["sub_rev"] + sums["onb_rev"]
    sums["cogs"] = sums["hosting"] + sums["support"] + sums["payfees"]
    sums["gross_profit"] = sums["revenue"] - sums["cogs"]
    sums["opex"] = sum(sums[k] for k in ("salaries", "marketing", "tools", "rent", "prof", "other"))
    sums["ebitda"] = sums["gross_profit"] - sums["opex"]
    budget = ms.engine(ms.drivers("Base"), ms.HIRES["Base"])
    scen = ms.engine(ms.drivers(scenario), ms.HIRES[scenario])
    bva_bad = 0
    for i, (key, label, typ, tot) in enumerate(bw.BVA_LINES):
        row = 6 + i
        ybud, yact = sum(budget[key][:n]), sums[key]
        var = (yact - ybud) if typ == "income" else (ybud - yact)
        fy_bud = sum(budget[key])
        fcst = sum(scen[key][n:])
        outlook = yact + fcst
        ovar = (outlook - fy_bud) if typ == "income" else (fy_bud - outlook)
        exp = {"B": ybud, "C": yact, "D": var, "G": fy_bud, "H": fcst, "I": outlook, "J": ovar}
        for col, val in exp.items():
            got = float(cell(dump, "Budget_vs_Actual", row, col))
            if abs(got - val) > 0.5:
                bva_bad += 1
                print(f"BvA MISMATCH {key} col {col}: excel {got} vs python {val}")
    print(f"Budget_vs_Actual cells mismatched ({scenario} outlook): {bva_bad}")

    # integrity checks sheet
    d = dump["Checks"]
    statuses = [row[3] for row in d["values"] if row and row[3] in ("PASS", "FAIL")]
    print(f"Checks sheet: {statuses.count('PASS')} PASS, {statuses.count('FAIL')} FAIL")
    overall = [row[1] for row in d["values"] if row and row[0] == "OVERALL STATUS"]
    print("Overall status cell:", overall)

    # no error values anywhere
    errs = 0
    for sheet, dd in dump.items():
        for r in dd["values"]:
            for v in r:
                if isinstance(v, str) and v in ERR_CODES and sheet != "ChartData":
                    errs += 1
                    print("ERROR VALUE in", sheet, v)
    print(f"error values outside ChartData: {errs}")
    return 0 if (fails == 0 and bva_bad == 0 and errs == 0 and statuses.count("FAIL") == 0) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "Base"))
