# Method and sources

## Why a local Python app plus an Excel export, and not a VBA workbook

- **Macros are blocked by default.** Microsoft Office blocks VBA macros in files that came from the internet, including email attachments. Users can no longer enable them with one click. A macro-driven calculator would break the first time you sent it to a prospect, and IT teams at engineering firms often lock macros down further. Source: Microsoft Learn, "Macros from the internet are blocked by default in Office".
- **Python can do things a spreadsheet cannot do cleanly.** It runs a Monte Carlo simulation, a sensitivity analysis, pricing solvers and automated tests, and it works offline during a screen-share.
- **Excel is still the buyer's language.** The export rebuilds the whole deterministic model as live formulas: 3,101 per workbook, with no macros. Finance can audit any number and change any input. `tests/verify_excel.py` recalculates each export in LibreOffice and checks that it matches the Python engine to within a billionth of a dollar.

## The model

**Three-point estimates.** Every uncertain input takes a low, likely and high value. People find ranges easier to give than single numbers. Hubbard's *How to Measure Anything* builds on the same idea: express uncertainty as a range, then calibrate it.

**Three cases.**
- Conservative takes the unfavourable end of every estimate at once. This follows the ROI Institute's guiding principles for Phillips ROI: choose the most conservative alternative, and adjust estimates for error.
- Expected takes the likely values.
- Optimistic takes the favourable end of every estimate.

You always quote the conservative case.

**Monthly cash flow over 36 months from signing.**
- Nothing is earned before go-live.
- After go-live, adoption climbs linearly to its steady-state share over the ramp period.
- The build fee is split between signing and delivery, and the audit fee can be credited against it.
- Payback is the first month in which cumulative net turns positive, interpolated within the month.
- NPV discounts the monthly flows. Forrester's TEI default discount rate is 10%.

**Cash versus capacity.** Freed hours count as cash only on three pathways:
- they remove overtime, valued at the overtime premium;
- they defer a hire, valued at loaded cost;
- they become billable work, valued at the billing rate.

Otherwise the hours are reported as capacity. Capacity is valued at loaded cost multiplied by a "share that turns into value", but only in the incl.-capacity view. That share follows Forrester TEI practice, which applies a 50% productivity recapture as standard. Here it is an explicit range.

**Soft costs count, soft benefits do not.** The cash view includes the client's own staff time for setup and for the owner's ongoing hours, valued at loaded cost. It does not include the capacity value.

**Time-and-materials trap.** When the task is billed hourly to end clients, every hour saved removes billed revenue. That loss is counted as a cost. On the billable pathway, refilled billable work is counted as a gain at the same rate. The net is hours × bill rate × (share converted to billable work − share billed).

**Loaded cost.** Salary × (1 + burden) / working hours. Canadian sources put employer payroll costs at about 10 to 15% of salary. With benefits included, the total runs about 1.15 to 1.4 times salary. The default burden is 25%, and 1,900 working hours is 52 weeks × 40 hours, less vacation and statutory holidays.

**Running costs.** Tokens per run × list price × USD to CAD rate, plus subscriptions. Claude prices were checked on 2026-11-19 against platform.claude.com/docs/en/about-claude/pricing:

| Model | Input, USD per M tokens | Output, USD per M tokens |
|---|---|---|
| Haiku 5.5 | 0.10 | 0.50 |
| Sonnet 5.5 | 2 | 10 |
| Opus 5.5 | 4 | 20 |
| GPT-5 mini (OpenAI) | 0.25 | 2.00 |

## Uncertainty

**Monte Carlo.** The simulation runs 3,000 draws in the app and 5,000 for the export. Each estimate is drawn from a PERT (beta) distribution bounded by its low and high values, with its mode at likely. The PERT mean is (low + 4 × likely + high) / 6. The seed is fixed, so the numbers stay still while you talk.

**Why both views.** The conservative case combines every pessimistic value at once, which Hubbard warns is very unlikely. The simulation shows how much of a floor it is: in the samples, the conservative case usually sits below every simulated outcome. Quote the conservative case. Use the probabilities to answer "how sure are you?"

**Caveat.** The estimates are drawn independently. In reality they often move together, for example low adoption and a long ramp, so read the simulated spread as a minimum.

**Sensitivity (tornado).** Each estimate is swung from its low to its high value while every other input stays at likely. The longest bar is the number to measure first.

## Why the baseline must be measured, not assumed

- **METR (2025)** ran a randomized controlled trial with experienced open-source developers. They forecast that AI would make them 24% faster. They were measured 19% slower, yet still believed they had been about 20% faster.
- **Dell'Acqua et al. (Organization Science; HBS/BCG field experiment with 758 consultants).** On tasks inside AI's capability, consultants with GPT-4 completed 12.2% more tasks, 25.1% faster, at higher quality. Outside that frontier they did worse.
- **Noy & Zhang (Science, 2023).** On mid-level writing tasks, time fell 40% and quality rose 18%.
- **MIT NANDA, "The GenAI Divide" (2025).** Most enterprise generative-AI pilots showed no measurable P&L impact. The biggest returns were in back-office automation.

The lesson for the calculator: gains are real but uneven, and self-reported time savings are unreliable. Time the task before the build and after it.

## Pricing guidance (internal, hidden in client view)

These are the project's heuristics, to be replaced by your own data:
- The build fee should be about one third of the value delivered in the first 90 days after go-live.
- The care plan should be 15 to 25% of the build fee per month.
- Payment is half up front and half on delivery.

The tool also solves for two numbers:
- the highest build fee that still passes the 6-month conservative payback rule;
- how much bigger the task would need to be to pass that rule at the fee you entered.

It flags a care plan that costs more than 30% of the conservative monthly benefit. On small builds, the 15 to 25% rule can sink payback on its own.
