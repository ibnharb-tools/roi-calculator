"""Hand-checkable tests for the ROI engine. Run: python -m unittest discover tests"""
import math
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import roi_engine as E  # noqa: E402


def simple(**kw):
    """10 runs/month, 120 -> 60 min, 100% adoption, instant go-live, no ramp,
    $50/h, hire pathway, 100% conversion: 10 h/month = $500/month cash.
    Only cost: $1,200 build fee paid up front."""
    c = E.blank_candidate("simple")
    flat = lambda x: [x, x, x]  # noqa: E731
    c.update(
        pathway="hire", billing_model="internal", error_type="cash",
        loaded_rate_override=50, runs_per_month=flat(10), minutes_before=flat(120),
        minutes_after=flat(60), adoption_pct=flat(100), ramp_months=flat(0), golive_weeks=flat(0),
        convert_pct=flat(100), errors_per_month=flat(0), error_reduction_pct=flat(0),
        cost_per_error=flat(0), other_value_month=flat(0), tokens_in=flat(0), tokens_out=flat(0),
        client_setup_hours=flat(0), owner_hours_month=flat(0), audit_fee=0, build_fee=1200,
        upfront_pct=100, care_plan_month=0, subscriptions_month=0, price_in_usd=0, price_out_usd=0,
    )
    c.update(kw)
    return c


class HandChecks(unittest.TestCase):
    def test_payback_and_year_one(self):
        r = E.evaluate(simple(), mc_runs=0)
        q = r["quoted"]["expected"]
        # month 1: +500 - 1200 = -700; m2: -200; m3: +300 -> 2 + 200/500
        self.assertAlmostEqual(q["payback"], 2.4, places=9)
        self.assertAlmostEqual(q["net"], 12 * 500 - 1200, places=6)
        self.assertAlmostEqual(r["cases"]["expected"]["hours_per_week_steady"], 10 * 12 / 52, places=9)

    def test_loaded_rate_from_salary(self):
        c, _ = E.normalise(simple(loaded_rate_override=0, salary=95000, burden_pct=20, paid_hours=1900))
        self.assertAlmostEqual(E.loaded_rate(c), 95000 * 1.2 / 1900, places=9)

    def test_ramp_and_golive(self):
        # go-live at 0.5 month, ramp 2 months
        self.assertEqual(E.ramp_factor(1, 0.5, 2), 0.5 * 0.25)
        self.assertEqual(E.ramp_factor(2, 0.5, 2), 1.0 * 0.75)
        self.assertEqual(E.ramp_factor(3, 0.5, 2), 1.0)
        self.assertEqual(E.ramp_factor(1, 1.5, 0), 0.0)

    def test_capacity_is_not_cash(self):
        r = E.evaluate(simple(pathway="capacity"), mc_runs=0)
        cash = r["cases"]["expected"]
        self.assertEqual(cash["y1_hard_benefit"], 0)
        self.assertAlmostEqual(cash["y1_soft_benefit"], 12 * 500, places=6)
        self.assertIsNone(r["quoted"]["expected"]["payback"])

    def test_overtime_premium(self):
        r = E.evaluate(simple(pathway="overtime", overtime_multiplier=1.5), mc_runs=0)
        self.assertAlmostEqual(r["cases"]["expected"]["y1_hard_benefit"], 12 * 750, places=6)

    def test_time_and_materials_revenue_at_risk(self):
        # all task hours billed at $100; freed hours refilled with billable work 100%
        c = simple(billing_model="time_and_materials", billable_share_pct=100, bill_rate=100,
                   pathway="billable")
        r = E.evaluate(c, mc_runs=0)
        e = r["cases"]["expected"]
        self.assertAlmostEqual(e["y1_rev_risk"], 12 * 10 * 100, places=6)
        # refilled revenue equals lost revenue, so net is just minus the fee
        self.assertAlmostEqual(r["quoted"]["expected"]["net"], -1200, places=6)

    def test_conservative_direction(self):
        c, _ = E.normalise(E.blank_candidate())
        v = E.case_values(c, "conservative")
        self.assertEqual(v["runs_per_month"], c["runs_per_month"][0])
        self.assertEqual(v["minutes_after"], c["minutes_after"][2])
        self.assertEqual(v["golive_weeks"], c["golive_weeks"][2])
        o = E.case_values(c, "optimistic")
        self.assertEqual(o["tokens_in"], c["tokens_in"][0])

    def test_case_ordering(self):
        r = E.evaluate(E.blank_candidate(), mc_runs=0)
        n = [r["quoted"][k]["net"] for k in E.CASES]
        self.assertLessEqual(n[0], n[1])
        self.assertLessEqual(n[1], n[2])

    def test_token_cost(self):
        c = simple(tokens_in=[1_000_000] * 3, tokens_out=[100_000] * 3, price_in_usd=2, price_out_usd=10,
                   fx_usd_cad=1.5)
        r = E.evaluate(c, mc_runs=0)
        self.assertAlmostEqual(r["cases"]["expected"]["token_cost_per_run"], (2 + 1) * 1.5, places=9)
        self.assertAlmostEqual(r["cases"]["expected"]["y1_running"], 12 * 10 * 4.5, places=6)

    def test_split_payment_and_audit_credit(self):
        c = simple(build_fee=2000, audit_fee=500, audit_credited=True, upfront_pct=50, golive_weeks=[13] * 3)
        r = E.monthly_model(*(lambda cc: (cc, E.case_values(cc, "expected")))(E.normalise(c)[0]))
        rows = r["rows"]
        # audit 500 + half of (2000-500) in month 1; other half in go-live month (13 wk = 3 months)
        self.assertAlmostEqual(rows[0]["one_time"], 500 + 750, places=9)
        self.assertAlmostEqual(rows[2]["one_time"], 750, places=9)

    def test_normalise_sorts_and_clamps(self):
        c, w = E.normalise({"runs_per_month": [10, 5, 7], "adoption_pct": [50, 90, 150]})
        self.assertEqual(c["runs_per_month"], [5, 7, 10])
        self.assertEqual(c["adoption_pct"][2], 100)
        self.assertTrue(w)

    def test_monte_carlo_is_reproducible(self):
        a = E.monte_carlo(E.normalise(E.blank_candidate())[0], n=500)
        b = E.monte_carlo(E.normalise(E.blank_candidate())[0], n=500)
        self.assertEqual(a["net_p50"], b["net_p50"])

    def test_pert_sample_bounds(self):
        import random
        rng = random.Random(1)
        xs = [E.pert_sample(rng, 2, 5, 20) for _ in range(5000)]
        self.assertTrue(min(xs) >= 2 and max(xs) <= 20)
        self.assertAlmostEqual(sum(xs) / len(xs), E.pert_mean(2, 5, 20), delta=0.2)

    def test_max_fee_solver(self):
        c, _ = E.normalise(simple())
        fee = E.max_fee_for_payback(c, 6)
        # 6 months x $500 = $3,000 supports a $3,000 up-front fee exactly
        self.assertAlmostEqual(fee, 3000, delta=1)

    def test_no_saving_flag(self):
        r = E.evaluate(simple(minutes_after=[200] * 3), mc_runs=0)
        self.assertTrue(any(f["level"] == "critical" for f in r["flags"]))
        self.assertIn("no time saving", r["call_sentence"])

    def test_call_sentence_has_no_em_dash(self):
        for sc in ("simple", "blank"):
            c = simple() if sc == "simple" else E.blank_candidate()
            r = E.evaluate(c, mc_runs=0)
            self.assertNotIn("—", r["call_sentence"])


if __name__ == "__main__":
    unittest.main()
