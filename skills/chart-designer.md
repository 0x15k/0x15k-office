---
name: Chart Designer & Visual Intelligence
description: Professional data visualization, chart type selection, DrawingML styling, color palettes, data-ink ratio optimization, zero-label suppression, and executive layout design for Microsoft Word, Excel, and PowerPoint.
hosts: [Word, Excel, PowerPoint]
alwaysActive: true
keywords: [chart, graph, visualization, drawingml, bar, column, pie, doughnut, line, palette, colors, data-ink, labels, aptos, typography, executive]
---

# Chart Designer & Visual Intelligence Skill

Adapted from the global Claude Office Skills and GenOffice ChartML specifications, this skill provides standard visualization intelligence for creating and transforming native charts in Microsoft Office.

---

## 1. Chart Selection Decision Tree

Select the optimal visual representation based on the communicative purpose and the underlying data dimensions:

```
What is the primary story of the data?
│
├── Comparison (Categories vs Metrics)
│   ├── Nominal Categories (≤ 7 items) → Horizontal Bar Chart (c:barChart barDir="bar")
│   ├── Sequential / Ordinal Categories → Vertical Column Chart (c:barChart barDir="col")
│   └── Multi-Series Comparison (e.g. Mitigated vs Vulnerable) → Stacked Bar Chart (c:grouping="stacked")
│
├── Part-to-Whole Composition
│   ├── Static Proportions (2 to 5 distinct slices) → Donut / Pie Chart (c:doughnutChart / c:pieChart)
│   │   └── CRITICAL RULE: If all items share the exact same status (e.g. 100% Mitigated), NEVER split a pie 50/50 without explicit status qualification!
│   └── Multi-Category Part-to-Whole → 100% Stacked Bar
│
├── Trend Over Time
│   ├── Continuous Time Series → Line Chart (c:lineChart with smooth curves)
│   └── Cumulative Volume → Area Chart (c:areaChart)
│
└── Distribution & Correlation
    ├── Two Numeric Variables → Scatter Plot (c:scatterChart)
    └── Metric vs Benchmark → Clustered Bar with target line
```

---

## 2. The Data-Ink Ratio & Label Hygiene Protocol

Edward Tufte's principle of Data-Ink Ratio is strictly enforced across all Office exhibits:

1. **Zero-Label Clutter Suppression**:
   - In Bar, Column, and Line charts, **NEVER display overlapping data labels on every category**.
   - If a series has 0 values (e.g. `0 Vulnerable` across all severities), displaying a large label `"VULNERABLE, CRÍTICA, 0"` creates dense visual noise.
   - **Default Rule**: `showDataLabels: false` for all bar and column charts unless the user explicitly requests *"etiquetas de datos"*, *"mostrar números"*, or *"labels"*.
   - When labels are enabled, apply the OpenXML format code `<c:formatCode>#;;</c:formatCode>` to automatically hide zero (0) values while displaying positive numbers cleanly.

2. **Axis Scale & Gridline Economy**:
   - The categorical axis (Y-axis for horizontal bars, X-axis for columns) must clearly show category names.
   - The value axis must use clean integer steps (e.g. `majorUnit="1"` for small integer counts like 0, 1, 2) rather than arbitrary decimals (0.2, 0.4, 0.6).
   - Major gridlines must use subtle slate borders (`#E2E8F0` or `<a:srgbClr val="E2E8F0"/>`), never heavy black lines.

---

## 3. Executive Typography & Styling Standards

Every native Word DrawingML chart must harmonize with the surrounding document typography:

- **Predominant Font Family**: Match `[DOCUMENT_STYLE_STATE]` or default to `Aptos` / `Aptos (Body)` / `Calibri`.
- **Type Sizes**:
  - **Chart Title**: 12pt Bold (`sz="1200"`), dark navy `#0F172A`. If the user asks for *"sin título"* or *"no title"*, emit `<c:autoTitleDeleted val="1"/>` to grant full vertical space to the plot area.
  - **Category Axis Labels**: 9pt Regular (`sz="900"`), charcoal `#334155`.
  - **Value Axis Ticks**: 8.5pt Regular (`sz="850"`), slate gray `#64748B`.
  - **Legend Entries**: 9pt Medium (`sz="900"`), positioned at the bottom (`<c:legendPos val="b"/>`) with `overlay="0"`.

---

## 4. Semantic Color Protocols & Brand Harmony

Colors must convey immediate meaning without ambiguity:

| Domain / Concept | Standard Hex | Semantic Purpose |
| :--- | :--- | :--- |
| **Mitigated / Subsanado / Resolved** | `#92D050` (or `#4ADE80`) | Positive resolution, safe condition |
| **Vulnerable / Open / Pendiente** | `#FF0000` (or `#EF4444`) | Active risk, requiring immediate action |
| **Crítica (Severity)** | `#501549` (or `#7F1D1D`) | Extreme risk level |
| **Alta (Severity)** | `#FF0000` (or `#DC2626`) | High risk level |
| **Media (Severity)** | `#FFC000` (or `#F59E0B`) | Medium risk level |
| **Baja (Severity)** | `#4EA72E` (or `#10B981`) | Low risk level |
| **Informativa (Severity)** | `#3B82F6` (or `#64748B`) | Informational note |
| **Corporate Primary** | `#002060` / `#0F172A` | Institutional authority |
| **Corporate Neutral / Grid** | `#E2E8F0` / `#94A3B8` | Structural frames, borders, axes |

*Rule*: If an active skill or customer profile defines brand colors, always pull those colors dynamically. Never hardcode client-specific color mappings into universal code.
