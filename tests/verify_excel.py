"""Export every candidate in scenarios/, recalculate with LibreOffice, and check
the workbook's live formulas reproduce the Python engine to the cent."""
import glob, json, os, subprocess, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import roi_engine as E
import excel_export as X
from openpyxl import load_workbook

RECALC = sys.argv[1]
fails = 0
tmp = tempfile.mkdtemp()
for path in sorted(glob.glob(os.path.join(os.path.dirname(__file__), "..", "scenarios", "*.json"))):
    sc = json.load(open(path))
    extra = []
    # stress variants: T&M with billable refill, cash errors, override rate, 0 ramp
    base = sc["candidates"][0]
    extra.append(dict(base, name="variant TM", billing_model="time_and_materials", billable_share_pct=70, pathway="billable", error_type="cash"))
    extra.append(dict(base, name="variant hire", pathway="hire", loaded_rate_override=88, ramp_months=[0, 0, 1], golive_weeks=[0, 2, 9], quote_basis="cash_plus_capacity", audit_credited=False, care_plan_month=400, subscriptions_month=50, other_value_month=[100, 300, 600]))
    for i, cand in enumerate(sc["candidates"] + extra):
        ev = E.evaluate(cand, mc_runs=200)
        wb = X.build_workbook(cand, sc.get("meta"), ev, E.compare(sc["candidates"]))
        out = os.path.join(tmp, f"{os.path.basename(path)}_{i}.xlsx")
        wb.save(out)
        res = json.loads(subprocess.run([sys.executable, RECALC, out, "60"], capture_output=True, text=True).stdout)
        if res.get("status") != "success":
            print("RECALC PROBLEM", out, res); fails += 1; continue
        ws = load_workbook(out, data_only=True)["Summary"]
        label_row = {ws.cell(row=r, column=1).value: r for r in range(1, 60)}
        checks = {
            "Hours freed per week at steady state": "hours_per_week_steady",
            "Hours freed in year one": "hours_year1",
            "Year-one cash benefit": "y1_hard_benefit",
            "Year-one capacity value (not cash)": "y1_soft_benefit",
            "Year-one cost": "y1_cost",
            "Year-one net, cash": "y1_net_cash",
            "Year-one net, incl. capacity value": "y1_net_all",
            "Payback, cash (months from signing)": "payback_cash_months",
            "Payback, incl. capacity (months)": "payback_all_months",
            "3-year NPV, cash": "npv3_cash",
            "3-year NPV, incl. capacity": "npv3_all",
        }
        worst = 0.0
        for label, key in checks.items():
            r = label_row[label]
            for j, case in enumerate(E.CASES):
                xv = ws.cell(row=r, column=2 + j).value
                pv = ev["cases"][case][key]
                if pv is None:
                    ok = isinstance(xv, str) and xv.startswith("Not within")
                    d = 0 if ok else 1e9
                else:
                    d = abs(float(xv) - pv) if not isinstance(xv, str) else 1e9
                worst = max(worst, d)
                if d > 0.01:
                    print(f"  MISMATCH {cand['name'][:30]} {case} {label}: excel={xv} python={pv}")
                    fails += 1
        print(f"{os.path.basename(path)} #{i} {cand['name'][:40]:40s} formulas={res['total_formulas']} worst diff={worst:.2e}")
print("FAILS:", fails)
sys.exit(1 if fails else 0)
