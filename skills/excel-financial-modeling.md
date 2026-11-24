---
name: Excel Financial Modeling & Dynamic Formulas
description: Designs institutional-grade financial models, dynamic array formulas (XLOOKUP, LET, SUMIFS), structured tables, and accounting formatting in Microsoft Excel.
hosts: [Excel]
alwaysActive: true
keywords: [excel, financial, model, formula, xlookup, let, sumifs, dcf, budget, forecast, balance sheet, p&l, cash flow, accounting, table, sheet]
---

## Core Principles

1. **Dynamic Native Formulas Over Static Values**:
   - NEVER calculate arithmetic results in JavaScript and hardcode static numbers into cells. Always emit dynamic Excel formulas via `set_formula` or `execute_office_js`.
   - Prefer modern dynamic array functions: `=XLOOKUP(lookup_val, lookup_range, return_range, [if_not_found])` instead of brittle `VLOOKUP`.
   - Use `=LET(...)` to declare clean intermediate calculation variables for readability and performance on complex calculations.
   - Use dynamic multi-criteria aggregations: `=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)` and `=AVERAGEIFS(...)`.
   - For percentage growth and variance: `=IFERROR((Actual - Budget) / Budget, 0)` with percentage formatting (`0.0%` or `+0.0%;-0.0%;0.0%`).

2. **Structured Table Automation**:
   - When creating datasets or transaction tables, always use `set_range_values` with `createTable: true` and specify a clean `tableName` (e.g. `RevenueModel`, `OperatingExpenses`).
   - For cells where formulas will be placed subsequently, insert empty strings `""` (never `null` or `0`) so the table structure initializes properly before formulas are bound.
   - Use structured table references where applicable: `=SUM(RevenueModel[Net Revenue])`.

3. **Wall Street / Institutional Formatting Standards**:
   - **Currency**: `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` or `$#,##0` for integer financial models.
   - **Accounting Zeroes**: Represent zeroes as dashes (`-`) rather than `$0.00` to prevent visual clutter in large matrices.
   - **Percentages**: `0.0%` for margins, returns, discount rates (WACC), and growth metrics.
   - **Multipliers**: `0.0"x"` for valuation multiples (EV/EBITDA, P/E, Debt/Equity).
   - **Header Styling**: Dark Navy blue background (`#1B365D`) with white bold text (`#FFFFFF`), or charcoal (`#2D3748`) with vertical center alignment.
   - **Summary / Total Rows**: Single top border (`thin`) and double bottom border (`double`), bold font, formula driven.

4. **Multi-Sheet Workbook Architecture**:
   - Structure complex models logically across worksheets:
     - `Assumptions`: WACC, tax rate, inflation, growth drivers.
     - `Income Statement`: Revenue build down to Net Income.
     - `Balance Sheet`: Assets, Liabilities, and Shareholder Equity.
     - `Cash Flow`: Operating, Investing, and Financing activities.
     - `Summary / Dashboard`: KPI cards, executive summary, and dynamic charts.
   - Reference external sheets cleanly: `=Assumptions!B4 * 'Income Statement'!C12`.

5. **Error Prevention & Resilience**:
   - Wrap division and lookup formulas in `=IFERROR(..., 0)` or `=IF(ISBLANK(...), "", ...)`.
   - Ensure range references never leave open-ended `#REF!` or `#VALUE!` hazards.
   - Always run `format_range` with `autoFit: true` on columns after inserting data to eliminate `###` overflow indicators.
