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

## On a call: the walkthrough

1. Tick **Client view** before you share your screen. It hides the file bar, the "Pricing and setup" tab and the internal coaching checks.
2. Click **New walkthrough**, type the company name, and go through **1 Questions** with the client, one step at a time:
   - The task
   - Time and volume
   - What the time costs
   - What goes wrong
   - The freed time
   - Rolling it out
   - With the automation (your estimate)
3. Each number takes one answer plus "How sure are you?". The answer sets how wide the range is:

   | Answer | Range around their number |
   |---|---|
   | Timed it | ±10% |
   | Pretty sure | 0.8x to 1.25x |
   | Rough guess | 0.6x to 1.5x |
   | No real idea | 0.4x to 2x |

   Click "adjust range" to set the low and high by hand.
4. Step 2 plays back the hours a month and the cost a year. Ask, "Does that sound right?" Their yes is the implication question answered in their own words.
5. The live strip shows the conservative numbers as you go. Untick "Show live numbers" to save the reveal for the end.
6. Open **2 Results** and read "In plain words". It always quotes the conservative case.
7. **3 Compare tasks** ranks every task you walked through (add more with **+ Task**).
8. Before or after the call, set fees, the AI model and token sizes in **Pricing and setup**. That tab is internal and shows the pricing guidance.

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
