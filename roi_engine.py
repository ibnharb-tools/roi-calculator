"""
ROI engine for Celestial Automation's automation-ROI calculator.

Method (see METHODOLOGY.md for sources):
  * Three-point estimates (low / likely / high) for every uncertain input.
  * Three deterministic cases. Conservative takes the unfavourable end of EVERY
    estimate (Phillips ROI guiding principle: "choose the most conservative
    alternative"). Expected takes the likely values. Optimistic takes the
    favourable end.
  * Monthly cash-flow model over 36 months from contract signing, with a
    go-live delay and a linear adoption ramp, so payback is measured in real
    calendar months, not "annual saving / cost".
  * Hard (cash) value is kept apart from soft (capacity) value. Freed hours are
    cash only when they remove overtime, defer a hire, or become billable work.
    Otherwise they are reported in hours, and optionally valued at loaded cost
    times a productivity-recapture share (Forrester TEI standard is 50%).
  * All costs are counted, including the client's own staff time (soft costs
    counted, soft benefits not, in the cash view).
  * Time-and-materials trap: billable revenue lost when billed work gets faster.
  * Monte Carlo over PERT (beta) distributions of the three-point inputs, plus a
    one-at-a-time sensitivity (tornado) to show which assumption to measure first.

Pure standard library. The Excel exporter reproduces the deterministic model in
live formulas; tests/test_engine.py checks the two agree.
"""
from __future__ import annotations

import copy
import math
import random
from typing import Any

HORIZON_MONTHS = 36
WEEKS_PER_MONTH = 52.0 / 12.0
CASES = ("conservative", "expected", "optimistic")

PRICES_VERIFIED_ON = "2026-11-19"
MODEL_PRESETS = [
    # USD per million tokens. Verified on PRICES_VERIFIED_ON against
    # platform.claude.com/docs/en/about-claude/pricing and OpenAI's price list
    # (via trackers). Re-check before quoting: prices change often.
    {"id": "none", "label": "No AI step (pure code / rules)", "in": 0.0, "out": 0.0},
    {"id": "claude-haiku-5-5", "label": "Claude Haiku 5.5", "in": 0.10, "out": 0.50},
    {"id": "claude-sonnet-5-5", "label": "Claude Sonnet 5.5", "in": 2.00, "out": 10.00},
    {"id": "claude-opus-5-5", "label": "Claude Opus 5.5", "in": 4.00, "out": 20.00},
    {"id": "gpt-5-mini", "label": "OpenAI GPT-5 mini", "in": 0.25, "out": 2.00},
    {"id": "custom", "label": "Custom price", "in": None, "out": None},
]

# ---------------------------------------------------------------------------
# Input schema. "dir" is the direction that is GOOD for the client:
#   +1  higher is better (benefit-side), conservative takes LOW
#   -1  higher is worse (cost-side),    conservative takes HIGH
# ---------------------------------------------------------------------------
TRIPLES: dict[str, dict[str, Any]] = {
    "runs_per_month":   {"label": "Runs per month", "unit": "runs/mo", "dir": +1},
    "minutes_before":   {"label": "Time per run today", "unit": "min", "dir": +1},
    "minutes_after":    {"label": "Time per run after (incl. human review)", "unit": "min", "dir": -1},
    "adoption_pct":     {"label": "Share of runs that go through the automation", "unit": "%", "dir": +1},
    "ramp_months":      {"label": "Months to reach full adoption after go-live", "unit": "months", "dir": -1},
    "golive_weeks":     {"label": "Weeks from signing to go-live", "unit": "weeks", "dir": -1},
    "convert_pct":      {"label": "Share of freed hours that turn into value", "unit": "%", "dir": +1},
    "errors_per_month": {"label": "Errors or rework events per month today", "unit": "errors/mo", "dir": +1},
    "error_reduction_pct": {"label": "Share of those errors the automation prevents", "unit": "%", "dir": +1},
    "cost_per_error":   {"label": "Cost per error", "unit": "CAD", "dir": +1},
    "other_value_month": {"label": "Other priced value per month at steady state", "unit": "CAD/mo", "dir": +1},
    "tokens_in":        {"label": "Input tokens per run", "unit": "tokens", "dir": -1},
    "tokens_out":       {"label": "Output tokens per run", "unit": "tokens", "dir": -1},
    "client_setup_hours": {"label": "Client staff hours for mapping, testing, training", "unit": "h", "dir": -1},
    "owner_hours_month": {"label": "Owner's ongoing hours per month", "unit": "h/mo", "dir": -1},
}

SINGLES: dict[str, dict[str, Any]] = {
    "salary":           {"label": "Annual salary of the person doing the task", "unit": "CAD/yr"},
    "burden_pct":       {"label": "Payroll burden and benefits on top of salary", "unit": "%"},
    "paid_hours":       {"label": "Working hours per year", "unit": "h/yr"},
    "loaded_rate_override": {"label": "Loaded hourly cost (override, 0 = use salary)", "unit": "CAD/h"},
    "overtime_multiplier": {"label": "Overtime premium multiplier", "unit": "x"},
    "bill_rate":        {"label": "Billing rate charged to clients", "unit": "CAD/h"},
    "billable_share_pct": {"label": "Share of task hours billed to clients (T&M only)", "unit": "%"},
    "audit_fee":        {"label": "Audit fee", "unit": "CAD"},
    "build_fee":        {"label": "Build fee (fixed)", "unit": "CAD"},
    "upfront_pct":      {"label": "Share of build fee paid up front", "unit": "%"},
    "care_plan_month":  {"label": "Care plan per month", "unit": "CAD/mo"},
    "subscriptions_month": {"label": "Tool and API subscriptions per month", "unit": "CAD/mo"},
    "price_in_usd":     {"label": "Model input price", "unit": "USD/M tokens"},
    "price_out_usd":    {"label": "Model output price", "unit": "USD/M tokens"},
    "fx_usd_cad":       {"label": "USD to CAD", "unit": "CAD per USD"},
    "discount_rate_pct": {"label": "Discount rate for 3-year NPV", "unit": "%/yr"},
}

CHOICES: dict[str, dict[str, Any]] = {
    "billing_model": {"label": "How the client bills this work",
                      "options": {"fixed_fee": "Fixed fee / lump sum",
                                  "time_and_materials": "Time and materials (hourly)",
                                  "internal": "Internal / overhead (not billed)"}},
    "pathway": {"label": "What happens to the freed hours",
                "options": {"capacity": "Absorbed into other work (capacity, not cash)",
                            "overtime": "Removes overtime (cash)",
                            "hire": "Defers or avoids a hire (cash)",
                            "billable": "Becomes billable work (cash)"}},
    "error_type": {"label": "What an error costs",
                   "options": {"cash": "Cash: write-offs, penalties, credits, lost work",
                               "rework": "Staff rework hours (follows the freed-hours rule)"}},
    "quote_basis": {"label": "Headline basis",
                    "options": {"cash": "Cash only (recommended)",
                                "cash_plus_capacity": "Cash plus capacity value"}},
}

FLAGS = ("audit_credited", "baseline_measured")


def blank_candidate(name: str = "New candidate") -> dict[str, Any]:
    """A template with neutral, clearly-marked placeholder values."""
    return {
        "name": name,
        "task_owner": "",
        "notes": "",
        "billing_model": "fixed_fee",
        "pathway": "capacity",
        "error_type": "rework",
        "quote_basis": "cash",
        "baseline_measured": False,
        "audit_credited": True,
        "model_preset": "claude-sonnet-5-5",
        "other_value_note": "",
        # labour
        "salary": 85000, "burden_pct": 25, "paid_hours": 1900, "loaded_rate_override": 0,
        "overtime_multiplier": 1.5, "bill_rate": 150, "billable_share_pct": 0,
        # three-point estimates [low, likely, high]
        "runs_per_month": [8, 12, 16],
        "minutes_before": [90, 120, 180],
        "minutes_after": [15, 25, 40],
        "adoption_pct": [60, 80, 95],
        "ramp_months": [1, 2, 3],
        "golive_weeks": [3, 4, 6],
        "convert_pct": [25, 50, 75],
        "errors_per_month": [0, 1, 2],
        "error_reduction_pct": [30, 50, 70],
        "cost_per_error": [100, 250, 500],
        "other_value_month": [0, 0, 0],
        "tokens_in": [20000, 30000, 50000],
        "tokens_out": [2000, 3000, 5000],
        "client_setup_hours": [6, 10, 16],
        "owner_hours_month": [1, 2, 4],
        # commercial
        "audit_fee": 750, "build_fee": 4500, "upfront_pct": 50, "care_plan_month": 0,
        "subscriptions_month": 0,
        "price_in_usd": 2.00, "price_out_usd": 10.00, "fx_usd_cad": 1.42,
        "discount_rate_pct": 10,
    }


# ---------------------------------------------------------------------------
# Validation / normalisation
# ---------------------------------------------------------------------------
def _num(x: Any, default: float = 0.0) -> float:
    try:
        v = float(x)
        if math.isnan(v) or math.isinf(v):
            return default
        return v
    except (TypeError, ValueError):
        return default


def normalise(c: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """Fill defaults, coerce types, sort each triple so low <= likely <= high.
    Returns (clean_candidate, list_of_input_warnings)."""
    base = blank_candidate()
    out = copy.deepcopy(base)
    warnings: list[str] = []
    for k, v in (c or {}).items():
        out[k] = v
    for k, meta in TRIPLES.items():
        t = out.get(k, base[k])
        if not isinstance(t, (list, tuple)) or len(t) != 3:
            t = [t, t, t] if not isinstance(t, (list, tuple)) else base[k]
        vals = [max(0.0, _num(x)) for x in t]
        if meta["unit"] == "%":
            vals = [min(100.0, x) for x in vals]
        if not (vals[0] <= vals[1] <= vals[2]):
            warnings.append(f"{meta['label']}: low, likely and high were out of order and have been sorted.")
            vals = sorted(vals)
        out[k] = vals
    for k in SINGLES:
        out[k] = max(0.0, _num(out.get(k), base[k]))
    for k in ("burden_pct", "billable_share_pct", "upfront_pct"):
        out[k] = min(100.0 if k != "burden_pct" else 200.0, out[k])
    if out["paid_hours"] <= 0:
        out["paid_hours"] = base["paid_hours"]
        warnings.append("Working hours per year must be above zero; reset to 1,900.")
    if out["fx_usd_cad"] <= 0:
        out["fx_usd_cad"] = base["fx_usd_cad"]
    for k, meta in CHOICES.items():
        if out.get(k) not in meta["options"]:
            out[k] = base[k]
    for k in FLAGS:
        out[k] = bool(out.get(k))
    out["name"] = str(out.get("name") or "Untitled candidate")[:120]
    return out, warnings


def loaded_rate(c: dict[str, Any]) -> float:
    if c["loaded_rate_override"] > 0:
        return c["loaded_rate_override"]
    return c["salary"] * (1 + c["burden_pct"] / 100.0) / c["paid_hours"]


def case_values(c: dict[str, Any], case: str) -> dict[str, float]:
    """Pick one value per three-point input for the given case."""
    v: dict[str, float] = {}
    for k, meta in TRIPLES.items():
        lo, mid, hi = c[k]
        if case == "expected":
            v[k] = mid
        elif case == "conservative":
            v[k] = lo if meta["dir"] > 0 else hi
        elif case == "optimistic":
            v[k] = hi if meta["dir"] > 0 else lo
        else:
            raise ValueError(case)
    return v


# ---------------------------------------------------------------------------
# Core monthly model. Every formula here is mirrored in excel_export.py.
# ---------------------------------------------------------------------------
def ramp_factor(t: int, G: float, R: float) -> float:
    """Share of steady-state adoption achieved in month t (1-based).
    G = go-live in months after signing, R = ramp length in months.
    Month is active for the part after go-live; adoption climbs linearly from
    0 at go-live to 1 at go-live + R (measured at month end)."""
    active = min(1.0, max(0.0, t - G))
    if R <= 0:
        return active
    return active * min(1.0, max(0.0, t - G) / R)


def monthly_model(c: dict[str, Any], v: dict[str, float]) -> dict[str, Any]:
    """Run the 36-month model for one set of point values v."""
    L = loaded_rate(c)
    G = v["golive_weeks"] * 12.0 / 52.0
    R = v["ramp_months"]
    golive_month = max(1, math.ceil(G - 1e-9))
    saved_h_per_run = max(0.0, v["minutes_before"] - v["minutes_after"]) / 60.0
    adopt = v["adoption_pct"] / 100.0
    conv = v["convert_pct"] / 100.0
    pathway = c["pathway"]
    if pathway == "overtime":
        cash_rate = L * c["overtime_multiplier"]
    elif pathway == "billable":
        cash_rate = c["bill_rate"]
    else:
        cash_rate = L
    tm = c["billing_model"] == "time_and_materials"
    billable_share = c["billable_share_pct"] / 100.0 if tm else 0.0
    token_cost_per_run = (v["tokens_in"] * c["price_in_usd"] + v["tokens_out"] * c["price_out_usd"]) / 1e6 * c["fx_usd_cad"]
    build_balance = c["build_fee"] - (c["audit_fee"] if c["audit_credited"] else 0.0)
    build_balance = max(0.0, build_balance)
    upfront = c["upfront_pct"] / 100.0
    rm = (1 + c["discount_rate_pct"] / 100.0) ** (1 / 12.0) - 1

    rows = []
    cum_cash = cum_all = 0.0
    payback_cash = payback_all = None
    for t in range(1, HORIZON_MONTHS + 1):
        f = ramp_factor(t, G, R)
        active = min(1.0, max(0.0, t - G))
        runs_auto = v["runs_per_month"] * adopt * f
        hours = runs_auto * saved_h_per_run
        # labour
        lab_hard = 0.0 if pathway == "capacity" else hours * conv * cash_rate
        lab_soft = hours * conv * L if pathway == "capacity" else 0.0
        # errors
        err_val = v["errors_per_month"] * (v["error_reduction_pct"] / 100.0) * adopt * f * v["cost_per_error"]
        if c["error_type"] == "cash":
            err_hard, err_soft = err_val, 0.0
        else:
            err_hard = 0.0 if pathway == "capacity" else err_val * conv
            err_soft = err_val * conv if pathway == "capacity" else 0.0
        other = v["other_value_month"] * f
        # T&M revenue at risk: billed hours that disappear
        rev_risk = hours * billable_share * c["bill_rate"]
        # costs
        one_time = 0.0
        if t == 1:
            one_time += c["audit_fee"] + build_balance * upfront + v["client_setup_hours"] * L
        if t == golive_month:
            one_time += build_balance * (1 - upfront)
        care = c["care_plan_month"] * active
        running = runs_auto * token_cost_per_run + c["subscriptions_month"] * active
        owner = v["owner_hours_month"] * L * active
        hard_benefit = lab_hard + err_hard + other
        soft_benefit = lab_soft + err_soft
        cost = one_time + care + running + owner + rev_risk
        net_cash = hard_benefit - cost
        net_all = net_cash + soft_benefit
        prev_cash, prev_all = cum_cash, cum_all
        cum_cash += net_cash
        cum_all += net_all
        if payback_cash is None and prev_cash <= 0 and cum_cash >= 0 and net_cash > 0:
            payback_cash = t - 1 + (-prev_cash) / net_cash
        if payback_all is None and prev_all <= 0 and cum_all >= 0 and net_all > 0:
            payback_all = t - 1 + (-prev_all) / net_all
        df = 1.0 / (1 + rm) ** t
        rows.append({
            "t": t, "factor": f, "runs_auto": runs_auto, "hours": hours,
            "lab_hard": lab_hard, "lab_soft": lab_soft, "err_hard": err_hard, "err_soft": err_soft,
            "other": other, "rev_risk": rev_risk, "one_time": one_time, "care": care,
            "running": running, "owner": owner, "hard_benefit": hard_benefit,
            "soft_benefit": soft_benefit, "cost": cost, "net_cash": net_cash, "net_all": net_all,
            "cum_cash": cum_cash, "cum_all": cum_all, "df": df,
        })

    def s(key: str, months: int = 12) -> float:
        return sum(r[key] for r in rows[:months])

    y1_cost = s("cost")
    y1_hard = s("hard_benefit")
    y1_soft = s("soft_benefit")
    res = {
        "loaded_rate": L,
        "golive_months": G,
        "saved_minutes_per_run": saved_h_per_run * 60,
        "hours_per_week_steady": v["runs_per_month"] * adopt * saved_h_per_run / WEEKS_PER_MONTH,
        "hours_year1": s("hours"),
        "y1_hard_benefit": y1_hard,
        "y1_soft_benefit": y1_soft,
        "y1_cost": y1_cost,
        "y1_one_time": s("one_time"),
        "y1_recurring": s("care") + s("running") + s("owner"),
        "y1_running": s("running"),
        "hours_3yr": s("hours", 36),
        "cost_per_hour_freed_y1": (y1_cost / s("hours")) if s("hours") > 0 else None,
        "cost_per_hour_freed_3yr": (s("cost", 36) / s("hours", 36)) if s("hours", 36) > 0 else None,
        "y1_rev_risk": s("rev_risk"),
        "y1_net_cash": y1_hard - y1_cost,
        "y1_net_all": y1_hard + y1_soft - y1_cost,
        "y1_roi_cash_pct": (y1_hard - y1_cost) / y1_cost * 100 if y1_cost > 0 else None,
        "y1_roi_all_pct": (y1_hard + y1_soft - y1_cost) / y1_cost * 100 if y1_cost > 0 else None,
        "payback_cash_months": payback_cash,
        "payback_all_months": payback_all,
        "npv3_cash": sum(r["net_cash"] * r["df"] for r in rows),
        "npv3_all": sum(r["net_all"] * r["df"] for r in rows),
        "total3_cash": s("net_cash", 36),
        "token_cost_per_run": token_cost_per_run,
        "rows": rows,
    }
    return res


def quoted(res: dict[str, Any], basis: str) -> dict[str, Any]:
    """Headline numbers on the chosen basis."""
    if basis == "cash_plus_capacity":
        return {"net": res["y1_net_all"], "payback": res["payback_all_months"],
                "roi": res["y1_roi_all_pct"], "npv": res["npv3_all"],
                "gain": res["y1_hard_benefit"] + res["y1_soft_benefit"]}
    return {"net": res["y1_net_cash"], "payback": res["payback_cash_months"],
            "roi": res["y1_roi_cash_pct"], "npv": res["npv3_cash"], "gain": res["y1_hard_benefit"]}


# ---------------------------------------------------------------------------
# Value in first 90 days after go-live (pricing heuristic: fee ~ 1/3 of it)
# ---------------------------------------------------------------------------
def value_first_90_days(c: dict[str, Any], v: dict[str, float], include_soft: bool) -> float:
    """Integrate gross benefit (before costs) from go-live to go-live + 3 months
    in 0.01-month steps, using the same ramp as the monthly model."""
    m = monthly_model(c, v)
    # monthly steady-state gross benefit rate at f = 1
    L = m["loaded_rate"]
    saved = m["saved_minutes_per_run"] / 60
    adopt = v["adoption_pct"] / 100
    conv = v["convert_pct"] / 100
    pathway = c["pathway"]
    rate = L * c["overtime_multiplier"] if pathway == "overtime" else (c["bill_rate"] if pathway == "billable" else L)
    hours_ss = v["runs_per_month"] * adopt * saved
    lab = hours_ss * conv * (L if pathway == "capacity" else rate)
    lab_counts = include_soft or pathway != "capacity"
    err = v["errors_per_month"] * v["error_reduction_pct"] / 100 * adopt * v["cost_per_error"]
    err_counts = c["error_type"] == "cash" or include_soft or pathway != "capacity"
    if c["error_type"] != "cash":
        err *= conv
    gross_ss = (lab if lab_counts else 0) + (err if err_counts else 0) + v["other_value_month"]
    R = v["ramp_months"]
    total, step, x = 0.0, 0.01, 0.0
    while x < 3.0 - 1e-12:
        mid = x + step / 2
        frac = 1.0 if R <= 0 else min(1.0, mid / R)
        total += gross_ss * frac * step
        x += step
    return total


# ---------------------------------------------------------------------------
# Monte Carlo with PERT (beta) distributions
# ---------------------------------------------------------------------------
def pert_sample(rng: random.Random, a: float, m: float, b: float) -> float:
    if b <= a:
        return a
    alpha = 1 + 4 * (m - a) / (b - a)
    beta = 1 + 4 * (b - m) / (b - a)
    return a + rng.betavariate(alpha, beta) * (b - a)


def pert_mean(a: float, m: float, b: float) -> float:
    return (a + 4 * m + b) / 6


def percentile(sorted_vals: list[float], p: float) -> float:
    if not sorted_vals:
        return 0.0
    k = (len(sorted_vals) - 1) * p
    lo, hi = math.floor(k), math.ceil(k)
    if lo == hi:
        return sorted_vals[int(k)]
    return sorted_vals[lo] * (hi - k) + sorted_vals[hi] * (k - lo)


def monte_carlo(c: dict[str, Any], n: int = 3000, seed: int = 20261119) -> dict[str, Any]:
    """Seeded so the numbers do not jitter while you talk on a call."""
    rng = random.Random(seed)
    basis = c["quote_basis"]
    nets, paybacks, npvs = [], [], []
    for _ in range(n):
        v = {k: pert_sample(rng, *c[k]) for k in TRIPLES}
        q = quoted(monthly_model(c, v), basis)
        nets.append(q["net"])
        paybacks.append(q["payback"] if q["payback"] is not None else math.inf)
        npvs.append(q["npv"])
    nets_s, npvs_s = sorted(nets), sorted(npvs)
    pb_s = sorted(paybacks)

    def pb_pct(p: float) -> float | None:
        val = pb_s[int(round((len(pb_s) - 1) * p))]
        return val if math.isfinite(val) else None

    # histogram of year-one net
    lo, hi = nets_s[0], nets_s[-1]
    bins = 24
    if hi - lo < 1e-9:
        hi = lo + 1
    width = (hi - lo) / bins
    counts = [0] * bins
    for x in nets:
        i = min(bins - 1, int((x - lo) / width))
        counts[i] += 1
    hist = [{"x0": lo + i * width, "x1": lo + (i + 1) * width, "count": counts[i]} for i in range(bins)]
    return {
        "n": n, "seed": seed, "basis": basis,
        "net_p10": percentile(nets_s, 0.10), "net_p50": percentile(nets_s, 0.50),
        "net_p90": percentile(nets_s, 0.90), "net_mean": sum(nets) / n,
        "npv_p10": percentile(npvs_s, 0.10), "npv_p50": percentile(npvs_s, 0.50),
        "npv_p90": percentile(npvs_s, 0.90),
        "p_net_positive": sum(1 for x in nets if x > 0) / n,
        "p_payback_6": sum(1 for p in paybacks if p <= 6) / n,
        "p_payback_12": sum(1 for p in paybacks if p <= 12) / n,
        "p_no_payback_36": sum(1 for p in paybacks if not math.isfinite(p)) / n,
        "payback_p50": pb_pct(0.5),
        "payback_p90": pb_pct(0.9),
        "hist": hist,
        "_nets_sorted": nets_s,
    }


# ---------------------------------------------------------------------------
# Sensitivity (tornado): swing each estimate low <-> high, others at likely
# ---------------------------------------------------------------------------
def tornado(c: dict[str, Any], top: int = 8) -> list[dict[str, Any]]:
    basis = c["quote_basis"]
    base_v = case_values(c, "expected")
    base = quoted(monthly_model(c, base_v), basis)["net"]
    out = []
    for k, meta in TRIPLES.items():
        lo, mid, hi = c[k]
        if hi == lo:
            continue
        vl = dict(base_v); vl[k] = lo
        vh = dict(base_v); vh[k] = hi
        nl = quoted(monthly_model(c, vl), basis)["net"]
        nh = quoted(monthly_model(c, vh), basis)["net"]
        swing = abs(nh - nl)
        if swing < 0.5:
            continue
        out.append({"key": k, "label": meta["label"], "unit": meta["unit"],
                    "low": lo, "high": hi, "net_at_low": nl, "net_at_high": nh,
                    "swing": swing, "base": base})
    out.sort(key=lambda r: r["swing"], reverse=True)
    return out[:top]


# ---------------------------------------------------------------------------
# Flags, coaching, pricing guidance, call sentence
# ---------------------------------------------------------------------------
def fmt_money(x: float | None) -> str:
    if x is None:
        return "n/a"
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1_000_000:
        return f"{sign}${x/1_000_000:.2f}M"
    if x >= 10_000:
        return f"{sign}${x/1000:.0f}k"
    if x >= 1000:
        return f"{sign}${x:,.0f}"
    return f"{sign}${x:,.0f}"


def fmt_months(x: float | None) -> str:
    if x is None:
        return "not within 36 months"
    if x < 1:
        return "under 1 month"
    return f"{x:.1f} months"


def _cons_payback(c: dict[str, Any]) -> float | None:
    return quoted(monthly_model(c, case_values(c, "conservative")), c["quote_basis"])["payback"]


def max_fee_for_payback(c: dict[str, Any], target: float = 6.0) -> float | None:
    """Largest build fee at which the conservative case still pays back within
    `target` months. None if it fails even at a zero build fee."""
    def ok(fee: float) -> bool:
        cc = dict(c); cc["build_fee"] = fee
        p = _cons_payback(cc)
        return p is not None and p <= target
    if not ok(0.0):
        return None
    lo, hi = 0.0, 1000.0
    while ok(hi) and hi < 1e7:
        lo, hi = hi, hi * 2
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if ok(mid) else (lo, mid)
    return lo


def volume_needed_for_payback(c: dict[str, Any], target: float = 6.0) -> dict[str, Any] | None:
    """How much bigger the task must be (scaling runs per month) for the
    conservative case to pay back within `target` months at the current fee."""
    def ok(m: float) -> bool:
        cc = dict(c); cc["runs_per_month"] = [x * m for x in c["runs_per_month"]]
        p = _cons_payback(cc)
        return p is not None and p <= target
    if ok(1.0):
        return {"multiplier": 1.0, "already_passes": True}
    if not ok(200.0):
        return None
    lo, hi = 1.0, 200.0
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if ok(mid) else (mid, hi)
    cc = dict(c); cc["runs_per_month"] = [x * hi for x in c["runs_per_month"]]
    hw = monthly_model(cc, case_values(cc, "conservative"))["hours_per_week_steady"]
    return {"multiplier": hi, "already_passes": False, "runs_per_month_low": cc["runs_per_month"][0],
            "hours_per_week_conservative": hw}


def pricing_guidance(c: dict[str, Any], cases: dict[str, Any]) -> dict[str, Any]:
    include_soft = c["quote_basis"] == "cash_plus_capacity" or c["pathway"] == "capacity"
    v90_cons = value_first_90_days(c, case_values(c, "conservative"), include_soft)
    v90_exp = value_first_90_days(c, case_values(c, "expected"), include_soft)
    fee_lo, fee_hi = v90_cons / 3, v90_exp / 3
    build = c["build_fee"]
    care_lo, care_hi = 0.15 * build, 0.25 * build
    notes = []
    if fee_hi > 0 and build < 0.6 * fee_lo:
        notes.append("Build fee is well under one third of conservative 90-day value. Check you are not pricing from nerves.")
    if fee_hi > 0 and build > 1.5 * fee_hi:
        notes.append("Build fee is well above one third of expected 90-day value. The price leans on results you cannot prove yet; shrink scope or offer a pilot.")
    if fee_hi <= 0:
        notes.append("No measurable value in the first 90 days on these inputs. Do not price this as scoped.")
    gross_month_cons = (cases["conservative"]["rows"][-1]["hard_benefit"] + cases["conservative"]["rows"][-1]["soft_benefit"])
    if c["care_plan_month"] > 0 and c["care_plan_month"] > 0.3 * max(1.0, gross_month_cons):
        notes.append(f"Care plan ({fmt_money(c['care_plan_month'])}/month) is more than 30% of the conservative monthly benefit at steady state ({fmt_money(gross_month_cons)}). On small builds the 15 to 25% of build fee rule can sink payback; price the care plan against monthly value instead.")
    if c["care_plan_month"] > 0 and not (care_lo * 0.8 <= c["care_plan_month"] <= care_hi * 1.2):
        notes.append(f"Care plan is outside the 15 to 25% of build fee heuristic ({fmt_money(care_lo)} to {fmt_money(care_hi)} per month).")
    max_fee = max_fee_for_payback(c)
    vol = volume_needed_for_payback(c)
    if max_fee is not None and build > max_fee * 1.001:
        notes.append(f"At this fee the conservative case misses 6-month payback. The most it supports is about {fmt_money(max_fee)}.")
    if vol and not vol["already_passes"]:
        notes.append(f"At this fee the task would need about {vol['multiplier']:.1f}x the volume ({vol['hours_per_week_conservative']:.1f} conservative hours a week) to pass 6-month payback.")
    if max_fee is None and vol is None:
        notes.append("No fee and no plausible volume makes the conservative case pay back in 6 months. Costs other than the build fee are too high for the value.")
    return {"max_fee_6mo": max_fee, "volume_for_6mo": vol,
            "value90_conservative": v90_cons, "value90_expected": v90_exp,
            "fee_low": fee_lo, "fee_high": fee_hi, "care_low": care_lo, "care_high": care_hi,
            "basis_note": "includes capacity value" if include_soft else "cash value only",
            "notes": notes}


def flags(c: dict[str, Any], cases: dict[str, Any], mc: dict[str, Any] | None) -> list[dict[str, str]]:
    """Coaching flags. 'audience': 'internal' ones are hidden in client view."""
    out: list[dict[str, str]] = []
    cons = cases["conservative"]
    exp = cases["expected"]
    qc = quoted(cons, c["quote_basis"])

    def add(level: str, text: str, audience: str = "internal") -> None:
        out.append({"level": level, "text": text, "audience": audience})

    if c["minutes_after"][1] >= c["minutes_before"][1]:
        add("critical", "Time after is not lower than time before. There is no time saving to sell on this task.", "client")
    if not c["baseline_measured"]:
        add("warning", "Baseline is an estimate. Time the task as it is done today (10 to 20 real runs) before quoting a firm number.", "client")
    pb = qc["payback"]
    if pb is None or pb > 6:
        add("serious", f"Conservative payback is {fmt_months(pb)}. Rule: over about 6 months, shrink the scope or walk away.")
    if exp["hours_per_week_steady"] < 5:
        add("warning", f"Expected saving is {exp['hours_per_week_steady']:.1f} hours a week. Below 5 the client will not feel it; look for a bigger bottleneck or bundle runs.")
    if c["billing_model"] == "time_and_materials":
        if c["billable_share_pct"] > 0:
            add("serious", f"Time-and-materials trap: {fmt_money(cons['y1_rev_risk'])} to {fmt_money(cases['optimistic']['y1_rev_risk'])} of billed revenue disappears in year one as billed hours shrink. Lead with fixed-fee margin, capacity to take more work, fewer errors and review cycles, and speed, not hours cut.")
        else:
            add("warning", "Client bills hourly. Confirm what share of this task's hours are billed to their clients; enter it, or the revenue at risk is missing.")
    if c["pathway"] == "capacity":
        add("info", "Freed hours are counted as capacity, not cash. Say it in hours. Cash only if they remove overtime, defer a hire, or become billable work.", "client")
    if c["convert_pct"][1] > 75 and c["pathway"] != "capacity":
        add("warning", "More than 75% of freed hours turning into cash is a strong claim. Forrester's TEI standard recaptures 50%. Be ready to defend it.")
    if not str(c.get("task_owner") or "").strip():
        add("warning", "No named owner yet. Every build needs one person on the client side who owns it.", "client")
    if c["owner_hours_month"][1] <= 0:
        add("warning", "No owner time entered. Every build needs a named owner on the client side; their hours are a real cost.")
    gross_cons = cons["y1_hard_benefit"] + cons["y1_soft_benefit"]
    if cons["y1_running"] > 0 and cons["y1_running"] > 0.25 * max(1.0, gross_cons):
        add("warning", f"Model and tool running costs ({fmt_money(cons['y1_running'])} in year one) exceed a quarter of the conservative benefit. Use a cheaper model for the bulk step or cut tokens per run.")
    if c["pathway"] == "capacity" and cons["cost_per_hour_freed_3yr"] is not None:
        add("info", f"Capacity framing: over three years each freed hour costs about {fmt_money(cons['cost_per_hour_freed_3yr'])} (conservative) against a loaded cost of {fmt_money(loaded_rate(c))} per hour.", "client")
    if mc:
        add("info", f"Simulation ({mc['n']:,} runs): {int(mc['p_net_positive']*100+0.5)}% chance year one is net positive; {int(mc['p_payback_6']*100+0.5)}% chance payback within 6 months.", "client")
    return out


def call_sentence(c: dict[str, Any], cases: dict[str, Any]) -> str:
    """Client-safe sentence in Kassem's voice: short, plain, no hype."""
    cons = cases["conservative"]
    q = quoted(cons, c["quote_basis"])
    cv = case_values(c, "conservative")
    hrs = cons["hours_per_week_steady"]
    adopt = cv["adoption_pct"]
    if c["minutes_after"][1] >= c["minutes_before"][1]:
        return "On these numbers there is no time saving to automate. I would not build this one."
    lead = (f"On conservative numbers, with {adopt:.0f}% of runs going through the automation, "
            f"the low end of your time-saving estimates, and every cost included, ")
    hours_txt = f"it frees about {hrs:.0f} hours a week" if hrs >= 1 else f"it frees about {hrs*60:.0f} minutes a week"
    if c["pathway"] == "capacity" and c["quote_basis"] == "cash":
        if q["net"] > 0 and q["payback"] is not None:
            body = (f"{hours_txt} once it is running. I am not counting those hours as cash. "
                    f"On cash alone it still pays back in {fmt_months(q['payback'])} and is {fmt_money(q['net'])} ahead in year one.")
        else:
            body = (f"{hours_txt} once it is running. I am not counting those hours as cash unless they cut overtime, "
                    f"delay a hire, or turn into billable work. That is the question for you: what would your team do with that time?")
    else:
        if q["payback"] is None or q["net"] <= 0:
            return (lead + f"this does not pay back within three years. "
                    "I would not build it as scoped. Let us find a bigger bottleneck or a smaller first step.")
        body = (f"{hours_txt}, pays back in {fmt_months(q['payback'])}, "
                f"and is {fmt_money(q['net'])} ahead after the first year.")
    tail = " The first step is timing the task as you do it today, which tightens these numbers." if not c["baseline_measured"] else ""
    return lead + body + tail


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------
def strip_rows(res: dict[str, Any], keep_rows: bool) -> dict[str, Any]:
    r = dict(res)
    if keep_rows:
        r["series"] = {
            "cum_cash": [row["cum_cash"] for row in res["rows"]],
            "cum_all": [row["cum_all"] for row in res["rows"]],
            "hours": [row["hours"] for row in res["rows"]],
        }
    r.pop("rows", None)
    return r


def evaluate(candidate: dict[str, Any], mc_runs: int = 3000, detail: bool = True) -> dict[str, Any]:
    c, input_warnings = normalise(candidate)
    cases_full = {case: monthly_model(c, case_values(c, case)) for case in CASES}
    out: dict[str, Any] = {
        "candidate": c,
        "input_warnings": input_warnings,
        "loaded_rate": loaded_rate(c),
        "case_values": {case: case_values(c, case) for case in CASES},
        "quoted": {case: quoted(cases_full[case], c["quote_basis"]) for case in CASES},
    }
    mc = monte_carlo(c, n=mc_runs) if (detail and mc_runs > 0) else None
    if mc:
        nets = mc.pop("_nets_sorted")
        qn = out["quoted"]["conservative"]["net"]
        below = sum(1 for x in nets if x < qn) / len(nets)
        mc["conservative_percentile"] = below
    out["cases"] = {case: strip_rows(cases_full[case], keep_rows=detail) for case in CASES}
    out["monte_carlo"] = mc
    out["tornado"] = tornado(c) if detail else []
    out["flags"] = flags(c, cases_full, mc)
    out["pricing"] = pricing_guidance(c, cases_full)
    out["call_sentence"] = call_sentence(c, cases_full)
    out["pert_means"] = {k: pert_mean(*c[k]) for k in TRIPLES}
    return out


def compare(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Audit view: rank candidates by conservative payback, then net."""
    rows = []
    for i, cand in enumerate(candidates):
        c, _ = normalise(cand)
        cons = monthly_model(c, case_values(c, "conservative"))
        exp = monthly_model(c, case_values(c, "expected"))
        qc, qe = quoted(cons, c["quote_basis"]), quoted(exp, c["quote_basis"])
        rows.append({
            "index": i, "name": c["name"],
            "payback_conservative": qc["payback"], "net_conservative": qc["net"],
            "net_expected": qe["net"], "hours_week_expected": exp["hours_per_week_steady"],
            "hours_week_conservative": cons["hours_per_week_steady"],
            "build_fee": c["build_fee"], "pathway": c["pathway"],
            "billing_model": c["billing_model"],
            "passes_6_month_rule": qc["payback"] is not None and qc["payback"] <= 6,
            "passes_5_hour_rule": exp["hours_per_week_steady"] >= 5,
        })
    rows.sort(key=lambda r: (r["payback_conservative"] if r["payback_conservative"] is not None else 1e9,
                             -r["net_conservative"]))
    for rank, r in enumerate(rows, 1):
        r["rank"] = rank
    return rows


def schema() -> dict[str, Any]:
    return {"triples": TRIPLES, "singles": SINGLES, "choices": CHOICES,
            "model_presets": MODEL_PRESETS, "prices_verified_on": PRICES_VERIFIED_ON,
            "horizon_months": HORIZON_MONTHS, "blank": blank_candidate()}
