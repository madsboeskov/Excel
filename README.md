# Capital IQ → DCF model builder (Excel VBA)

A VBA macro that reads the historical financials you pull from S&P Capital IQ
into Excel and builds a formula-driven unlevered DCF model on a new sheet
called **DCF**.

* Works with Capital IQ **Excel exports** (Financials → Export to Excel) and
  with data **pasted** from the Capital IQ website or pulled with the
  Capital IQ Excel plug-in, as long as the sheet has line-item labels in one
  column and fiscal years across the columns.
* The data can be on one sheet or split across several (e.g. separate
  Income Statement, Balance Sheet and Cash Flow sheets). Items are matched by
  fiscal year, not by column position.
* Every number on the DCF sheet is a live formula. Historical columns link back
  to the Capital IQ cells (green), so refreshing the data in the same layout
  updates the model. Projection drivers are yellow cells you can overwrite.

## Also in this repository

| Folder | Purpose |
| --- | --- |
| `survey/` | A small website with a questionnaire that automatically e-mails a confirmation to whoever fills it in, and stores the answers for download as CSV. See [`survey/README.md`](survey/README.md) (in Danish). |

## Files

| File | Purpose |
| --- | --- |
| `src/CapIQ_DCF.bas` | The VBA module to import into Excel |
| `sample/CapIQ_Sample.xlsx` | Example Capital IQ-style data to try the macro on |
| `sample/CapIQ_Sample_with_DCF.xlsx` | The same data with the DCF sheet the macro generates |

## Installation

1. Open the workbook that holds your Capital IQ data (or `sample/CapIQ_Sample.xlsx`).
2. Press **Alt+F11** to open the VBA editor (Mac: Tools → Macro → Visual Basic Editor).
3. **File → Import File…** and select `src/CapIQ_DCF.bas`.
4. Save the workbook as **.xlsm** (macro-enabled) if you want to keep the macro in it.
   To have it available in every workbook, import it into your `PERSONAL.XLSB` instead.

## Usage

1. Press **Alt+F8**, choose **BuildDCF** and click **Run**.
2. Enter the sheet name(s) that hold the Capital IQ data, separated by commas, or
   `*` to search every sheet. The default is `CapIQ Data` if that sheet exists.
3. The macro creates (or replaces) the **DCF** sheet and shows which Capital IQ
   lines it used. Items it could not find are listed in red at the bottom of the
   DCF sheet and use defaults or 0. Check those first.

You can also call it from your own code without dialogs:

```vba
Dim summary As String
summary = BuildDCFFromSheets(ActiveWorkbook, "Income Statement, Balance Sheet, Cash Flow")
```

## What the macro looks for

**Year header.** The row, within the first 60, with the most fiscal-year
headers, for example `FY2023`, `2023`, `2023A`, `Dec-31-2023`,
`12 months Dec-31-2023` or a real date. LTM, NTM, YTD, quarterly and estimate
columns (`2025E`, `Est.`) are skipped. The model uses the last 5 historical
years that have revenue.

**Line items.** Matched on the label column, ignoring case and punctuation.
Exact matches win over "starts with" matches, which win over "contains" matches.
Ratio rows (%, growth, margin, per share, …) are ignored.

| Model item | Capital IQ labels searched (in priority order) |
| --- | --- |
| Revenue* | Total Revenue, Revenue, Net Sales, Sales, Turnover, … |
| EBIT* | EBIT, Operating Income, Operating Profit, … |
| D&A | Depreciation & Amort., Total; Depreciation & Amort.; D&A, … |
| Pre-tax income | EBT, Incl. Unusual Items; EBT, Excl. Unusual Items; EBT; Pre-Tax Income, … |
| Income tax | Income Tax Expense, Income Taxes, Provision for Income Taxes, … |
| Capex | Capital Expenditure, Capex, Purchase of Property, Plant & Equipment, … |
| Change in NWC | Change in Net Working Capital, Change in Working Capital, … |
| Cash | Total Cash & ST Investments, Cash And Equivalents, … |
| Debt | Total Debt, Total Borrowings |
| Minority interest | Minority Interest, Non-Controlling Interest, … |
| Preferred | Total Preferred Equity, Preferred Stock |
| Shares | Weighted Avg. Diluted Shares Out., Total Shares Out. on Filing Date, … |
| Beta | Beta, Levered Beta, 5 Year Beta |
| Share price | Share Price, Stock Price, Last Close Price, Price, … |

\* required. Everything else falls back to a default or 0.

If your template uses different captions, add them to `ItemSynonyms` in the module.

**Sign conventions.** Capital IQ shows expenses such as D&A, tax and capex as
negative numbers. The macro uses absolute values for D&A, tax and capex, and
reads "Change in Net Working Capital" as a cash-flow item, so a negative number
means working capital increased. Balance-sheet items and shares come from the
latest historical year. Beta and share price come from the right-most number on
their row.

## The generated model

**Key assumptions.** Share price, diluted shares, risk-free rate (4.25%),
equity risk premium (5.5%), beta, pre-tax cost of debt (6%), tax rate
(average historical effective rate), market-value capital-structure weights,
WACC (CAPM), terminal growth (2.5%), exit EV/EBITDA multiple (10.0x), the
terminal value method (1 = Gordon growth, 2 = exit multiple, 3 = average) and a
mid-year-convention switch.

**Free cash flow projection.** Historical years plus 5 projected years:

* Revenue growth fades in a straight line from the historical CAGR to the terminal growth rate.
* EBIT margin, D&A % of revenue and capex % of revenue default to their historical averages.
* NWC investment defaults to the historical median as a % of the change in revenue.
* Unlevered FCF = EBIT − taxes on EBIT + D&A − capex − increase in NWC.
* Each year shows a discount period, a discount factor and the PV.

**Valuation.** Sum of PV of FCF + PV of terminal value = enterprise value, then
− debt + cash − minority interest − preferred = equity value, divided by
diluted shares to get the implied value per share and the upside against the
current price. Sanity checks: implied EV/LTM EBITDA, the exit multiple implied
by the Gordon growth value, and the perpetual growth implied by the exit
multiple.

**Sensitivity tables.** Implied value per share for WACC × terminal growth and
for WACC × exit multiple. Both are plain formulas, not Excel data tables, so
they recalculate immediately.

**Mapping log.** Which Capital IQ label and cell each model item came from.

Colour code: **yellow** = assumption you can overwrite, **blue** = hard-coded
input, **green** = link to Capital IQ data, **black** = formula.

## Configuration

Constants at the top of the module:

| Constant | Default | Meaning |
| --- | --- | --- |
| `DEFAULT_DATA_SHEET` | `CapIQ Data` | Sheet suggested in the prompt |
| `OUTPUT_SHEET` | `DCF` | Name of the generated sheet |
| `PROJ_YEARS` | 5 | Number of projected years |
| `MAX_HIST_YEARS` | 5 | Historical years shown |
| `DEF_*` | see above | Default WACC / terminal-value inputs |
| `SENS_*_STEP` | 0.5% / 0.25% / 1.0x | Sensitivity table step sizes |

## Notes and limitations

* All figures stay in the units of the Capital IQ data, which is usually
  millions. Shares must be in the same unit (Capital IQ reports them in
  millions) for the per-share value to be correct.
* The model is a starting point. Review the yellow drivers, especially the
  growth fade, margins and NWC, before using the output.
* Formulas are written with `Range.Formula`, which always takes English
  function names and commas, so the macro also works with non-English Excel
  locales such as Danish.
* Works in Excel for Windows and Mac. It uses no Windows-only libraries.
