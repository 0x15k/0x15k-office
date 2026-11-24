---
name: Excel Data Analyst & Native Visualization
description: Cleanses dirty datasets, computes exploratory statistical metrics, generates pivot summaries, and builds native Excel charts (Column, Line, Pie, Bar) via Office.js.
hosts: [Excel]
alwaysActive: true
keywords: [excel, data, analysis, analyst, chart, graph, plot, pivot, statistics, outliers, clean, deduplicate, filter, sort]
---

## Core Capabilities

1. **Native Excel Chart Generation**:
   When the user asks to visualize data or build a chart in Excel, use `execute_office_js` to call the native `worksheet.charts.add(...)` API:
   ```javascript
   await Excel.run(async (context) => {
     const sheet = context.workbook.worksheets.getActiveWorksheet();
     // Define data range (e.g., A1:D10 including headers)
     const dataRange = sheet.getRange("A1:D10");
     // Add chart: ColumnClustered, Line, BarClustered, Pie, or Scatter
     const chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.auto);
     chart.title.text = "Quarterly Revenue Breakdown (2025–2026)";
     chart.legend.position = Excel.ChartLegendPosition.right;
     chart.legend.format.fill.setSolidColor("#F8F9FA");
     // Position the chart neatly next to or below the dataset
     chart.top = 220;
     chart.left = 20;
     chart.height = 320;
     chart.width = 560;
     await context.sync();
   });
   ```
   - Chart Types supported:
     - `Excel.ChartType.columnClustered` (comparison across categories)
     - `Excel.ChartType.line` / `lineMarkers` (time series and trends)
     - `Excel.ChartType.barClustered` (ranking and horizontal comparisons)
     - `Excel.ChartType.pie` / `doughnut` (share of total, <= 6 categories)
     - `Excel.ChartType.areaStacked` (cumulative volume over time)

2. **Automated Data Cleansing & Normalization**:
   - Trim leading/trailing whitespace and non-printable characters (`=TRIM(CLEAN(cell))`).
   - Standardize dates to ISO `YYYY-MM-DD` or regional standard format.
   - Standardize case for categories (`=PROPER(cell)` or `=UPPER(cell)`).
   - Identify and highlight duplicate keys or missing mandatory values using conditional formatting.

3. **Exploratory Data Analysis (EDA) Summary Blocks**:
   - When asked to analyze a dataset, generate a dedicated summary table with key statistical metrics:
     - Count of Records (`=COUNTA(...)`)
     - Unique Values (`=COUNTA(UNIQUE(...))`)
     - Arithmetic Mean (`=AVERAGE(...)`)
     - Median (`=MEDIAN(...)`)
     - Standard Deviation (`=STDEV.S(...)`)
     - Min / Max Range (`=MIN(...)` to `=MAX(...)`)
     - Interquartile Range (`=QUARTILE.INC(..., 3) - QUARTILE.INC(..., 1)`)

4. **Dynamic Aggregation Tables (Pivot Patterns)**:
   - Construct summary aggregation tables using dynamic formulas:
     - Distinct categories: `=SORT(UNIQUE(Data[Category]))`
     - Category counts: `=COUNTIF(Data[Category], F2)`
     - Category volume: `=SUMIF(Data[Category], F2, Data[Amount])`
     - Share of Total: `=G2 / SUM($G$2:$G$10)` (formatted as `0.0%`)
