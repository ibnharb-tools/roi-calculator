"""
Macro-free Excel export. Rebuilds the engine's deterministic model in live
formulas so a prospect's finance team can audit every number and change any
input. Monte Carlo results are pasted as static values (Excel cannot rerun
them without macros), and are labelled as such.

Internal pricing guidance (max fee, fee heuristics, coaching flags) is NOT
exported: this file is meant to be sent to the client.
"""
from __future__ import annotations

import datetime as dt
from typing import Any

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import roi_engine as E

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
GREEN = Font(name=FONT, color="008000")
BLACK = Font(name=FONT, color="000000")
BOLD = Font(name=FONT, bold=True)
TITLE = Font(name=FONT, bold=True, size=14, color="1B2140")
MUTED = Font(name=FONT, italic=True, color="666666", size=9)
HEAD_FILL = PatternFill("solid", fgColor="1B2140")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
SECTION_FILL = PatternFill("solid", fgColor="E8EAF2")
CONS_FILL = PatternFill("solid", fgColor="DCE9F9")
THIN = Side(style="thin", color="C3C2B7")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0%;(0%);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM0 = '#,##0;(#,##0);"-"'
FACT = '0.00'

PCT_SINGLES = {"burden_pct", "billable_share_pct", "upfront_pct", "discount_rate_pct"}

SINGLE_NAMES = {
    "salary": "Salary", "burden_pct": "BurdenPct", "paid_hours": "PaidHours",
    "loaded_rate_override": "RateOverride", "overtime_multiplier": "OTMult", "bill_rate": "BillRate",
    "billable_share_pct": "BillableShare", "audit_fee": "AuditFee", "build_fee": "BuildFee",
    "upfront_pct": "UpfrontPct", "care_plan_month": "CarePlan", "subscriptions_month": "Subscriptions",
    "price_in_usd": "PriceIn", "price_out_usd": "PriceOut", "fx_usd_cad": "FX",
    "discount_rate_pct": "DiscountRate",
}
CHOICE_NAMES = {"billing_model": "BillingModel", "pathway": "Pathway", "error_type": "ErrorType",
                "quote_basis": "QuoteBasis"}
SINGLE_FMT = {"salary": CUR, "paid_hours": NUM0, "loaded_rate_override": CUR2, "overtime_multiplier": '0.00"x"',
              "bill_rate": CUR2, "audit_fee": CUR, "build_fee": CUR, "care_plan_month": CUR,
              "subscriptions_month": CUR, "price_in_usd": CUR2, "price_out_usd": CUR2, "fx_usd_cad": '0.0000'}
TRIPLE_FMT = {"%": PCT, "CAD": CUR, "CAD/mo": CUR, "tokens": NUM0}

MONTH_COLS = [
    ("t", "Month", NUM0),
    ("factor", "Adoption ramp factor", FACT),
    ("active", "Share of month live", FACT),
    ("runs_auto", "Runs automated", NUM1),
    ("hours", "Hours freed", NUM1),
    ("lab_hard", "Labour value, cash", CUR),
    ("lab_soft", "Labour value, capacity", CUR),
    ("err_raw", "Error cost avoided (gross)", CUR),
    ("err_hard", "Error value, cash", CUR),
    ("err_soft", "Error value, capacity", CUR),
    ("other", "Other priced value", CUR),
    ("rev_risk", "Billed revenue at risk (T&M)", CUR),
    ("one_time", "One-time costs", CUR),
    ("care", "Care plan", CUR),
    ("running", "Model and tool running cost", CUR),
    ("owner", "Owner time", CUR),
    ("hard_benefit", "Cash benefit", CUR),
    ("soft_benefit", "Capacity value (not cash)", CUR),
    ("cost", "Total cost", CUR),
    ("net_cash", "Net cash", CUR),
    ("cum_cash", "Cumulative net cash", CUR),
    ("net_all", "Net incl. capacity", CUR),
    ("cum_all", "Cumulative incl. capacity", CUR),
    ("pb_cash", "Payback marker (cash)", NUM1),
    ("pb_all", "Payback marker (incl. capacity)", NUM1),
    ("df", "Discount factor", '0.0000'),
    ("pv_cash", "PV net cash", CUR),
    ("pv_all", "PV net incl. capacity", CUR),
]
COL = {key: get_column_letter(i + 1) for i, (key, _, _) in enumerate(MONTH_COLS)}

CASE_SHEETS = {"conservative": "Conservative", "expected": "Expected", "optimistic": "Optimistic"}
CASE_COL_IN_INPUTS = {"conservative": "G", "expected": "H", "optimistic": "I"}


def _add_name(wb: Workbook, name: str, ref: str) -> None:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _head_row(ws, row: int, labels: list[str], start_col: int = 1) -> None:
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")


def _inputs_sheet(wb: Workbook, c: dict[str, Any], meta: dict[str, Any]) -> dict[str, int]:
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"] = f"Inputs: {c['name']}"
    ws["A1"].font = TITLE
    ws["A2"] = ("Blue cells are inputs; change any of them and every sheet recalculates. "
                "Yellow cells are the three-point estimates that matter most. Percentages are stored as fractions (0.8 = 80%).")
    ws["A2"].font = MUTED
    ws.merge_cells("A2:J2")
    ws.row_dimensions[2].height = 30
    ws["A2"].alignment = Alignment(wrap_text=True)

    r = 4
    ws.cell(row=r, column=1, value="Settings").font = BOLD
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    dv_rows = {}
    for key, meta_c in E.CHOICES.items():
        ws.cell(row=r, column=1, value=meta_c["label"]).font = BLACK
        cell = ws.cell(row=r, column=2, value=c[key])
        cell.font = BLUE
        cell.border = BOX
        ws.cell(row=r, column=3, value="Options: " + "; ".join(f"{k} = {v}" for k, v in meta_c["options"].items())).font = MUTED
        _add_name(wb, CHOICE_NAMES[key], f"Inputs!$B${r}")
        dv = DataValidation(type="list", formula1='"' + ",".join(meta_c["options"].keys()) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(cell)
        dv_rows[key] = r
        r += 1
    ws.cell(row=r, column=1, value="Audit fee credited toward the build").font = BLACK
    cell = ws.cell(row=r, column=2, value="Yes" if c["audit_credited"] else "No")
    cell.font = BLUE
    cell.border = BOX
    dv = DataValidation(type="list", formula1='"Yes,No"')
    ws.add_data_validation(dv)
    dv.add(cell)
    _add_name(wb, "AuditCredited", f"Inputs!$B${r}")
    r += 2

    ws.cell(row=r, column=1, value="Rates, fees and prices").font = BOLD
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    for key, meta_s in E.SINGLES.items():
        ws.cell(row=r, column=1, value=meta_s["label"]).font = BLACK
        val = c[key] / 100.0 if key in PCT_SINGLES else c[key]
        cell = ws.cell(row=r, column=2, value=val)
        cell.font = BLUE
        cell.border = BOX
        cell.number_format = PCT if key in PCT_SINGLES else SINGLE_FMT.get(key, NUM1)
        ws.cell(row=r, column=3, value=meta_s["unit"]).font = MUTED
        _add_name(wb, SINGLE_NAMES[key], f"Inputs!$B${r}")
        if key in ("price_in_usd", "price_out_usd"):
            cell.comment = Comment(f"List price per million tokens, checked {E.PRICES_VERIFIED_ON}. Prices change; re-check before relying on them.", "Celestial Automation")
        if key == "fx_usd_cad":
            cell.comment = Comment(f"Approximate mid-November 2026 rate (Bank of Canada about 1.40 in late October; market about 1.42 to 1.43 mid-November). Update to today's rate.", "Celestial Automation")
        if key == "burden_pct":
            cell.comment = Comment("Employer CPP, EI, health tax, vacation and benefits. Canadian estimates run about 15 to 40% of salary depending on benefits.", "Celestial Automation")
        r += 1
    ws.cell(row=r, column=1, value="Loaded hourly cost (calculated)").font = BOLD
    cell = ws.cell(row=r, column=2, value="=IF(RateOverride>0,RateOverride,Salary*(1+BurdenPct)/PaidHours)")
    cell.font = BLACK
    cell.number_format = CUR2
    cell.border = BOX
    ws.cell(row=r, column=3, value="Salary x (1 + burden) / working hours, unless overridden").font = MUTED
    _add_name(wb, "LoadedRate", f"Inputs!$B${r}")
    r += 2

    ws.cell(row=r, column=1, value="Three-point estimates").font = BOLD
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.cell(row=r, column=3, value="Conservative takes the unfavourable end of every estimate; optimistic the favourable end. Direction +1 means higher is better for you.").font = MUTED
    r += 1
    _head_row(ws, r, ["Estimate", "Unit", "Low", "Likely", "High", "Direction", "Conservative", "Expected", "Optimistic", "PERT mean"])
    r += 1
    triple_rows: dict[str, int] = {}
    for key, meta_t in E.TRIPLES.items():
        lo, mid, hi = c[key]
        scale = 0.01 if meta_t["unit"] == "%" else 1.0
        fmt = TRIPLE_FMT.get(meta_t["unit"], NUM1)
        ws.cell(row=r, column=1, value=meta_t["label"]).font = BLACK
        ws.cell(row=r, column=2, value="%" if meta_t["unit"] == "%" else meta_t["unit"]).font = MUTED
        for j, val in enumerate((lo, mid, hi)):
            cell = ws.cell(row=r, column=3 + j, value=val * scale)
            cell.font = BLUE
            cell.fill = KEY_FILL
            cell.border = BOX
            cell.number_format = fmt
        d = ws.cell(row=r, column=6, value=meta_t["dir"])
        d.font = BLUE
        d.alignment = Alignment(horizontal="center")
        for col, f in (("G", f"=IF(F{r}=1,C{r},E{r})"), ("H", f"=D{r}"), ("I", f"=IF(F{r}=1,E{r},C{r})"),
                       ("J", f"=(C{r}+4*D{r}+E{r})/6")):
            cell = ws[f"{col}{r}"]
            cell.value = f
            cell.font = BLACK
            cell.number_format = fmt
            if col == "G":
                cell.fill = CONS_FILL
        triple_rows[key] = r
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Baseline").font = BOLD
    ws.cell(row=r, column=2, value="Measured" if c["baseline_measured"] else "Estimated, not yet timed").font = BLUE
    r += 1
    if c.get("other_value_note"):
        ws.cell(row=r, column=1, value="What the 'other priced value' is").font = BOLD
        ws.cell(row=r, column=2, value=c["other_value_note"]).font = BLUE
        r += 1
    if c.get("notes"):
        ws.cell(row=r, column=1, value="Notes").font = BOLD
        ws.cell(row=r, column=2, value=c["notes"]).font = BLACK
        r += 1
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 24
    for col in "CDEFGHIJ":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["C"].width = 16
    ws.freeze_panes = "B4"
    return triple_rows


def _case_sheet(wb: Workbook, case: str, triple_rows: dict[str, int]) -> dict[str, str]:
    ws = wb.create_sheet(CASE_SHEETS[case])
    ws["A1"] = f"{CASE_SHEETS[case]} case: month-by-month model (36 months from signing)"
    ws["A1"].font = TITLE
    ws["A2"] = "Every cell is a formula. Values in the block below come from the Inputs sheet."
    ws["A2"].font = MUTED
    incol = CASE_COL_IN_INPUTS[case]
    ref: dict[str, str] = {}
    r = 4
    _head_row(ws, r - 1, ["Case value", "Value"])
    for key, meta_t in E.TRIPLES.items():
        ws.cell(row=r, column=1, value=meta_t["label"])
        cell = ws.cell(row=r, column=2, value=f"=Inputs!{incol}{triple_rows[key]}")
        cell.font = GREEN
        cell.number_format = TRIPLE_FMT.get(meta_t["unit"], NUM1)
        ref[key] = f"$B${r}"
        r += 1
    derived = [
        ("G", "Go-live, months after signing", f"={ref['golive_weeks']}*12/52", NUM1),
        ("GM", "Go-live month (second build payment)", None, NUM0),
        ("saved_h", "Hours saved per automated run", f"=MAX(0,{ref['minutes_before']}-{ref['minutes_after']})/60", '0.00'),
        ("cash_rate", "Value per converted hour (cash pathways)",
         '=IF(Pathway="overtime",LoadedRate*OTMult,IF(Pathway="billable",BillRate,LoadedRate))', CUR2),
        ("share", "Billed share of task hours (T&M only)", '=IF(BillingModel="time_and_materials",BillableShare,0)', PCT),
        ("tok", "Model cost per run (CAD)", f"=({ref['tokens_in']}*PriceIn+{ref['tokens_out']}*PriceOut)/1000000*FX", '$#,##0.0000'),
        ("build_bal", "Build fee after audit credit", '=MAX(0,BuildFee-IF(AuditCredited="Yes",AuditFee,0))', CUR),
        ("rm", "Monthly discount rate", "=(1+DiscountRate)^(1/12)-1", '0.000%'),
    ]
    for key, label, formula, fmt in derived:
        ws.cell(row=r, column=1, value=label)
        ref[key] = f"$B${r}"
        if key == "GM":
            formula = f"=MAX(1,ROUNDUP({ref['G']}-0.000000001,0))"
        cell = ws.cell(row=r, column=2, value=formula)
        cell.number_format = fmt
        r += 1

    head = r + 1
    _head_row(ws, head, [lab for _, lab, _ in MONTH_COLS])
    ws.row_dimensions[head].height = 45
    first = head + 1
    last = first + E.HORIZON_MONTHS - 1
    R = ref
    for i in range(E.HORIZON_MONTHS):
        row = first + i
        p = row - 1
        A = f"{COL['t']}{row}"
        f = {
            "t": i + 1,
            "factor": f"=MIN(1,MAX(0,{A}-{R['G']}))*IF({R['ramp_months']}<=0,1,MIN(1,MAX(0,{A}-{R['G']})/{R['ramp_months']}))",
            "active": f"=MIN(1,MAX(0,{A}-{R['G']}))",
            "runs_auto": f"={R['runs_per_month']}*{R['adoption_pct']}*{COL['factor']}{row}",
            "hours": f"={COL['runs_auto']}{row}*{R['saved_h']}",
            "lab_hard": f'=IF(Pathway="capacity",0,{COL["hours"]}{row}*{R["convert_pct"]}*{R["cash_rate"]})',
            "lab_soft": f'=IF(Pathway="capacity",{COL["hours"]}{row}*{R["convert_pct"]}*LoadedRate,0)',
            "err_raw": f"={R['errors_per_month']}*{R['error_reduction_pct']}*{R['adoption_pct']}*{COL['factor']}{row}*{R['cost_per_error']}",
            "err_hard": f'=IF(ErrorType="cash",{COL["err_raw"]}{row},IF(Pathway="capacity",0,{COL["err_raw"]}{row}*{R["convert_pct"]}))',
            "err_soft": f'=IF(ErrorType="cash",0,IF(Pathway="capacity",{COL["err_raw"]}{row}*{R["convert_pct"]},0))',
            "other": f"={R['other_value_month']}*{COL['factor']}{row}",
            "rev_risk": f"={COL['hours']}{row}*{R['share']}*BillRate",
            "one_time": (f"=IF({A}=1,AuditFee+{R['build_bal']}*UpfrontPct+{R['client_setup_hours']}*LoadedRate,0)"
                         f"+IF({A}={R['GM']},{R['build_bal']}*(1-UpfrontPct),0)"),
            "care": f"=CarePlan*{COL['active']}{row}",
            "running": f"={COL['runs_auto']}{row}*{R['tok']}+Subscriptions*{COL['active']}{row}",
            "owner": f"={R['owner_hours_month']}*LoadedRate*{COL['active']}{row}",
            "hard_benefit": f"={COL['lab_hard']}{row}+{COL['err_hard']}{row}+{COL['other']}{row}",
            "soft_benefit": f"={COL['lab_soft']}{row}+{COL['err_soft']}{row}",
            "cost": f"={COL['one_time']}{row}+{COL['care']}{row}+{COL['running']}{row}+{COL['owner']}{row}+{COL['rev_risk']}{row}",
            "net_cash": f"={COL['hard_benefit']}{row}-{COL['cost']}{row}",
            "cum_cash": f"={COL['net_cash']}{row}" if i == 0 else f"={COL['cum_cash']}{p}+{COL['net_cash']}{row}",
            "net_all": f"={COL['net_cash']}{row}+{COL['soft_benefit']}{row}",
            "cum_all": f"={COL['net_all']}{row}" if i == 0 else f"={COL['cum_all']}{p}+{COL['net_all']}{row}",
            "df": f"=1/(1+{R['rm']})^{A}",
            "pv_cash": f"={COL['net_cash']}{row}*{COL['df']}{row}",
            "pv_all": f"={COL['net_all']}{row}*{COL['df']}{row}",
        }
        for kind, cumk, netk in (("pb_cash", "cum_cash", "net_cash"), ("pb_all", "cum_all", "net_all")):
            prev = "0" if i == 0 else f"{COL[cumk]}{p}"
            f[kind] = (f'=IF(AND({prev}<=0,{COL[cumk]}{row}>=0,{COL[netk]}{row}>0),'
                       f'{A}-1+(-{prev})/{COL[netk]}{row},"")')
        for key, _, fmt in MONTH_COLS:
            cell = ws[f"{COL[key]}{row}"]
            cell.value = f[key]
            cell.number_format = fmt
            cell.font = BLACK
            if i == 11:
                cell.border = Border(bottom=Side(style="medium", color="1B2140"))
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 14
    for key, _, _ in MONTH_COLS[2:]:
        ws.column_dimensions[COL[key]].width = 13
    ws.freeze_panes = ws[f"C{first}"]
    rng = {key: f"'{CASE_SHEETS[case]}'!${COL[key]}${first}:${COL[key]}${last}" for key, _, _ in MONTH_COLS}
    y1 = {key: f"'{CASE_SHEETS[case]}'!${COL[key]}${first}:${COL[key]}${first + 11}" for key, _, _ in MONTH_COLS}
    out = {"_sheet": CASE_SHEETS[case], "_first": first}
    out.update({f"all:{k}": v for k, v in rng.items()})
    out.update({f"y1:{k}": v for k, v in y1.items()})
    out.update({f"ref:{k}": f"'{CASE_SHEETS[case]}'!{v}" for k, v in ref.items()})
    return out


def _summary_sheet(wb: Workbook, c: dict[str, Any], meta: dict[str, Any], refs: dict[str, dict[str, str]],
                   ev: dict[str, Any]) -> dict[str, int]:
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = f"Automation ROI: {c['name']}"
    ws["A1"].font = Font(name=FONT, bold=True, size=16, color="1B2140")
    ws["A2"] = (f"Prepared by {meta.get('prepared_by') or 'Celestial Automation'} for "
                f"{meta.get('client') or 'the client'}, {dt.date.today().isoformat()}. All figures CAD.")
    ws["A2"].font = MUTED
    ws["A3"] = ("Quote the conservative case. It takes the unfavourable end of every estimate at once, "
                "so it is a floor, not a forecast.")
    ws["A3"].font = MUTED
    _head_row(ws, 5, ["Measure", "Conservative", "Expected", "Optimistic", "Note"])
    lines = [
        ("hpw", "Hours freed per week at steady state",
         lambda R: f"={R['ref:runs_per_month']}*{R['ref:adoption_pct']}*{R['ref:saved_h']}/(52/12)", NUM1,
         "Runs x adoption x hours saved per run"),
        ("h1", "Hours freed in year one", lambda R: f"=SUM({R['y1:hours']})", NUM0, "Includes go-live delay and adoption ramp"),
        ("gain", "Year-one cash benefit", lambda R: f"=SUM({R['y1:hard_benefit']})", CUR,
         "Counts freed hours as cash only for overtime, hire or billable pathways"),
        ("soft", "Year-one capacity value (not cash)", lambda R: f"=SUM({R['y1:soft_benefit']})", CUR,
         "Freed hours x share that turns into value x loaded cost"),
        ("cost", "Year-one cost", lambda R: f"=SUM({R['y1:cost']})", CUR,
         "Fees, running costs, client staff time, revenue at risk"),
        ("one", "   of which one-time", lambda R: f"=SUM({R['y1:one_time']})", CUR, "Audit, build fee, client setup hours"),
        ("rec", "   of which recurring", lambda R: f"=SUM({R['y1:care']})+SUM({R['y1:running']})+SUM({R['y1:owner']})", CUR,
         "Care plan, models and tools, owner time"),
        ("risk", "   of which billed revenue at risk", lambda R: f"=SUM({R['y1:rev_risk']})", CUR,
         "Time-and-materials only"),
        ("net", "Year-one net, cash", lambda R: None, CUR, "Cash benefit minus cost"),
        ("netall", "Year-one net, incl. capacity value", lambda R: None, CUR, "Adds capacity value"),
        ("roi", "Year-one ROI, cash", lambda R: None, PCT, "Net / cost"),
        ("pb", "Payback, cash (months from signing)",
         lambda R: f'=IF(COUNT({R["all:pb_cash"]})=0,"Not within 36 months",MIN({R["all:pb_cash"]}))', NUM1, ""),
        ("pball", "Payback, incl. capacity (months)",
         lambda R: f'=IF(COUNT({R["all:pb_all"]})=0,"Not within 36 months",MIN({R["all:pb_all"]}))', NUM1, ""),
        ("npv", "3-year NPV, cash", lambda R: f"=SUM({R['all:pv_cash']})", CUR, "Discounted monthly at the input rate"),
        ("npvall", "3-year NPV, incl. capacity", lambda R: f"=SUM({R['all:pv_all']})", CUR, ""),
        ("cph", "Cost per freed hour over 3 years",
         lambda R: f"=IF(SUM({R['all:hours']})>0,SUM({R['all:cost']})/SUM({R['all:hours']}),0)", CUR2,
         "Compare with the loaded hourly cost on Inputs"),
    ]
    rows: dict[str, int] = {}
    r = 6
    for key, label, fn, fmt, note in lines:
        rows[key] = r
        ws.cell(row=r, column=1, value=label).font = BOLD if not label.startswith("   ") else BLACK
        for j, case in enumerate(E.CASES):
            col = get_column_letter(2 + j)
            R = refs[case]
            if key == "net":
                formula = f"={col}{rows['gain']}-{col}{rows['cost']}"
            elif key == "netall":
                formula = f"={col}{rows['gain']}+{col}{rows['soft']}-{col}{rows['cost']}"
            elif key == "roi":
                formula = f"=IF({col}{rows['cost']}>0,{col}{rows['net']}/{col}{rows['cost']},0)"
            else:
                formula = fn(R)
            cell = ws.cell(row=r, column=2 + j, value=formula)
            cell.number_format = fmt
            cell.font = GREEN if "!" in formula else BLACK
            if case == "conservative":
                cell.fill = CONS_FILL
        ws.cell(row=r, column=5, value=note).font = MUTED
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Headline on the chosen basis").font = BOLD
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    for label, cash_key, all_key, fmt in (("Year-one net", "net", "netall", CUR), ("Payback (months)", "pb", "pball", NUM1)):
        ws.cell(row=r, column=1, value=label)
        for j in range(3):
            col = get_column_letter(2 + j)
            cell = ws.cell(row=r, column=2 + j,
                           value=f'=IF(QuoteBasis="cash",{col}{rows[cash_key]},{col}{rows[all_key]})')
            cell.number_format = fmt
            if j == 0:
                cell.fill = CONS_FILL
                cell.font = BOLD
        r += 1
    ws.cell(row=r, column=1, value="Basis").font = MUTED
    ws.cell(row=r, column=2, value='=IF(QuoteBasis="cash","Cash only","Cash plus capacity value")').font = MUTED
    r += 2

    mc = ev.get("monte_carlo")
    if mc:
        ws.cell(row=r, column=1, value="Uncertainty (simulated, static values)").font = BOLD
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.cell(row=r, column=2, value=(f"{mc['n']:,} draws from PERT distributions of every three-point estimate, "
                                       "run in the calculator app. These cells do NOT update when you edit inputs here.")).font = MUTED
        r += 1
        for label, val, fmt in (
            ("Chance year one is net positive", mc["p_net_positive"], PCT),
            ("Chance of payback within 6 months", mc["p_payback_6"], PCT),
            ("Chance of payback within 12 months", mc["p_payback_12"], PCT),
            ("Year-one net, 10th percentile", mc["net_p10"], CUR),
            ("Year-one net, median", mc["net_p50"], CUR),
            ("Year-one net, 90th percentile", mc["net_p90"], CUR),
            ("Share of simulated outcomes below the conservative case", mc.get("conservative_percentile", 0), '0.0%'),
        ):
            ws.cell(row=r, column=1, value=label)
            cell = ws.cell(row=r, column=2, value=val)
            cell.number_format = fmt
            cell.font = BLACK
            r += 1
        r += 1
    ws.cell(row=r, column=1, value="What we would say on a call").font = BOLD
    r += 1
    ws.cell(row=r, column=1, value=ev.get("call_sentence", "")).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.row_dimensions[r].height = 62
    r += 2
    ws.cell(row=r, column=1, value="How to read this").font = BOLD
    r += 1
    for line in (
        "Hours freed count as cash only when they remove overtime, defer a hire, or become billable work. Otherwise they are capacity, shown in hours.",
        "Costs include your own staff time for setup and the named owner's monthly hours, valued at loaded cost.",
        "A measured baseline beats any estimate. The first step of any engagement is timing the task as it is done today.",
        "Sheets 'Conservative', 'Expected' and 'Optimistic' show every month. 'Method' lists the sources behind the approach.",
    ):
        ws.cell(row=r, column=1, value="- " + line).font = MUTED
        r += 1
    ws.column_dimensions["A"].width = 46
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["E"].width = 62
    ws.freeze_panes = "B6"
    return rows


def _compare_sheet(wb: Workbook, compare_rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet("Candidates")
    ws["A1"] = "Audit ranking: candidates by conservative payback"
    ws["A1"].font = TITLE
    ws["A2"] = "Static values from the calculator at export time. Each candidate has its own workbook with live formulas."
    ws["A2"].font = MUTED
    _head_row(ws, 4, ["Rank", "Candidate", "Payback, conservative (months)", "Year-one net, conservative",
                      "Year-one net, expected", "Hours a week, expected", "Build fee"])
    r = 5
    for row in compare_rows:
        vals = [row["rank"], row["name"],
                row["payback_conservative"] if row["payback_conservative"] is not None else "Not within 36 months",
                row["net_conservative"], row["net_expected"], row["hours_week_expected"], row["build_fee"]]
        fmts = [NUM0, None, NUM1, CUR, CUR, NUM1, CUR]
        for j, (v, fm) in enumerate(zip(vals, fmts)):
            cell = ws.cell(row=r, column=1 + j, value=v)
            if fm:
                cell.number_format = fm
        r += 1
    ws.column_dimensions["B"].width = 46
    for col in "CDEFG":
        ws.column_dimensions[col].width = 18


def _method_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Method")
    ws["A1"] = "Method and sources"
    ws["A1"].font = TITLE
    lines = [
        ("Three cases from three-point estimates",
         "Each uncertain input has a low, likely and high value. The conservative case takes the unfavourable end of every one, following the ROI Institute's guiding principle to choose the most conservative alternative. Phillips & Phillips, ROI Institute."),
        ("Monthly cash flow, not annual averages",
         "Benefits start at go-live and ramp up linearly, so payback reflects real calendar months from signing."),
        ("Productivity recapture",
         "Not every freed hour becomes value. Forrester's Total Economic Impact studies apply a 50% recapture rate as their standard; here it is an explicit input with a range."),
        ("Cash versus capacity",
         "Freed hours are cash only if they remove overtime, defer a hire, or become billable work. Otherwise they are reported as capacity."),
        ("Time-and-materials billing",
         "Where the task is billed hourly to end clients, hours saved reduce billed revenue. That loss is counted as a cost."),
        ("Uncertainty",
         "The app samples each three-point estimate from a PERT (beta) distribution, mean = (low + 4 x likely + high) / 6, thousands of times (Monte Carlo). Hubbard, How to Measure Anything, notes that combining every pessimistic value into one case is far less likely than any single value, so the simulation shows how much of a floor the conservative case is."),
        ("Why measure the baseline",
         "Self-reported time savings from AI tools can be badly wrong. In a 2025 randomized trial (METR), experienced developers expected AI to make them 24% faster and were measured 19% slower. Field experiments find large gains on tasks inside AI's capability and worse results outside it (Dell'Acqua et al., Organization Science; Noy & Zhang, Science 2023). So the baseline and the result are timed, not assumed."),
        ("Model prices",
         f"List prices in USD per million tokens, checked {E.PRICES_VERIFIED_ON}; converted to CAD at the FX input. Prices change often."),
    ]
    r = 3
    for head, body in lines:
        ws.cell(row=r, column=1, value=head).font = BOLD
        cell = ws.cell(row=r, column=2, value=body)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = max(30, 15 * (len(body) // 95 + 1))
        r += 1
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 110


def build_workbook(candidate: dict[str, Any], meta: dict[str, Any] | None = None,
                   evaluation: dict[str, Any] | None = None,
                   compare_rows: list[dict[str, Any]] | None = None) -> Workbook:
    meta = meta or {}
    c, _ = E.normalise(candidate)
    ev = evaluation or E.evaluate(c)
    wb = Workbook()
    triple_rows = _inputs_sheet(wb, c, meta)
    refs = {case: _case_sheet(wb, case, triple_rows) for case in E.CASES}
    _summary_sheet(wb, c, meta, refs, ev)
    if compare_rows and len(compare_rows) > 1:
        _compare_sheet(wb, compare_rows)
    _method_sheet(wb)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.font is None or cell.font.name != FONT:
                    f = cell.font
                    cell.font = Font(name=FONT, bold=f.bold, italic=f.italic, color=f.color, size=f.size)
    wb.active = 0
    return wb
