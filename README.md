# Celestial Automation ROI calculator

A local tool for working out automation ROI live on a call. It runs only on your machine, needs no internet, and exports a macro-free Excel workbook you can send to the prospect.

## Start it

1. Install Python 3.9 or newer (python.org). On Windows, tick "Add Python to PATH".
2. Once only: `pip install openpyxl`. This is needed only for the Excel export.
3. Double-click `Start ROI calculator.bat` on Windows or `start_roi_calculator.command` on Mac, or run:

```
python roi_calculator.py
```

Your browser opens at http://127.0.0.1:8765. Press Ctrl+C in the terminal window to stop it.

## On a call

The calculator has three steps, in plain English:

1. **Your process.** The task, the hours a week, the people, their hourly cost, what mistakes cost, and how sure the client is of those numbers.
2. **The automation.** How much time it saves, team uptake, build time and cost, the monthly cost, who looks after it, what the freed time is used for, and whether the work is billed by the hour.
3. **Results.** The first-year savings, break-even, hours back each week, 3-year savings, three charts, and plain-language advice. The Cautious case is the default and the one you quote.

The menu (☰) holds Start a new one, Open, Save, and "Show my pricing notes", which is internal and off by default. The moon/sun button switches theme.

The earlier detailed version, with seven steps and audit ranking, is still available at http://127.0.0.1:8765/static/advanced.html.

## During an audit

Add one tab per candidate task (+ Add). **Compare and rank** sorts the tasks by conservative payback and checks each one against the 6-month payback rule and the 5-hours-a-week rule. Recommend the top task as the first build.

## Files

| Path | What it is |
|---|---|
| `roi_calculator.py` | The local web server. Listens on 127.0.0.1 only. |
| `roi_engine.py` | The model: three cases, monthly cash flow, Monte Carlo, sensitivity, pricing solvers. Uses only the standard library. |
| `excel_export.py` | Rebuilds the model as live Excel formulas. Uses no macros. |
| `static/index.html` | The interface. |
| `scenarios/` | Saved scenarios as JSON. `_autosave.json` restores your last session. |
| `exports/` | Every Excel file you export is also saved here. |
| `tests/` | `python -m unittest discover tests` runs the engine tests. `tests/verify_excel.py <path to recalc.py>` checks the Excel formulas against the engine. |
| `METHODOLOGY.md` | The method, the reasons behind it, and the sources. |

## Two sample scenarios

- **demo_evidence_bank_1**: your compressed-air audit tool, using the times from the evidence bank. The salary and fees are placeholders, so use it for rehearsal only.
- **illustrative_three_candidates**: three made-up tasks at an engineering consultancy. Use it to practise the audit ranking and the time-and-materials conversation.

## Look and feel

The colours, fonts (Sora, IBM Plex Sans, IBM Plex Mono), light and dark themes, buttons, labels and orbit line art are copied from the Celestial Automation website (2026-11-19 redesign) so the two match. The theme toggle in the header remembers your choice, using the same setting name as the site. The fonts are bundled in `static/fonts` (SIL Open Font License), so the tool still works offline. If the site's palette changes, update the `:root` block at the top of `static/index.html`.

## Keep current

Model prices were checked on 2026-11-19 against Anthropic's pricing page and OpenAI's list prices. The USD to CAD rate defaults to 1.42. Both change, so re-check them before quoting. The presets are at the top of `roi_engine.py` in `MODEL_PRESETS`.
