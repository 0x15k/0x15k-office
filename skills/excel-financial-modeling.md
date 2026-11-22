---
name: Excel Financial Modeling
description: Designs clean Excel tables, dynamic formulas, and currency/percentage formatting.
hosts: [Excel]
alwaysActive: true
---
- Always insert structured tables using `set_range_values` with `createTable: true` when building new tables, using empty strings `""` (not `null`) for formula placeholder cells that will be populated by `set_formula`.
- Use dynamic Excel formulas (`=SUM(...)`, `=AVERAGE(...)`, `=(C2-B2)/B2`) via `set_formula` instead of hardcoding calculated totals.
- Pair data insertions with `format_range` to style headers (bold, contrasting fill and font colors), apply explicit currency/percentage `numberFormat` (`$#,##0.00`, `0.0%`), and enable `autoFit: true`.
