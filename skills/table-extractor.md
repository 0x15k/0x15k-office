---
name: Table Extractor & Tabular Data Intelligence
description: Precise extraction, normalization, and structured aggregation of data tables from Word documents and PDFs for analytical processing and chart generation.
hosts: [Word, Excel]
alwaysActive: true
keywords: [table, extraction, parse, rows, columns, headers, metrics, aggregation, chart data, tabular, fidelity]
---

# Table Extractor & Tabular Data Intelligence Skill

Adapted from the Claude Office Skills `table-extractor` module, this skill standardizes the extraction, normalization, and aggregation of tabular data from Office documents.

---

## 1. Table Detection & Header Identification

When reading documents or prompts containing markdown tables or Word tables:
1. **Header Identification**:
   - Locate the primary column headers (Row 0).
   - Classify columns into:
     - **Identifier / Code**: `Código`, `ID`, `Ref`, `Ticket`.
     - **Categorical / Dimensional**: `Vulnerabilidad`, `Category`, `Severidad`, `Criticidad`, `Department`, `Product`.
     - **Metric / Numeric**: `Score`, `CVSS`, `Amount`, `Count`, `Quantity`, `Percentage`.
     - **Status / State**: `Estado`, `Status`, `Condición`, `Resolution`.

2. **Dimension Mapping**:
   - For chart generation, identify which column serves as the **Category Dimension** (e.g. `Criticidad` or `Category`) and which serves as the **Series Dimension** (e.g. `Estado` or `Period`).
   - Group rows by category and series, aggregating row counts or summing numeric metrics strictly from the extracted rows without inventing data.

---

## 2. Table-to-Chart Metric Translation

When an extracted table serves as the source of truth for an executive visual exhibit:

1. **Category Normalization**:
   - Map ordinal or categorical levels into clean, standardized categories:
     - Example: If a table lists `Media` and `Baja`, the complete scale for a security audit spans: `["CRÍTICA", "ALTA", "MEDIA", "BAJA", "INFORMATIVA"]`.
   - Categories present in the scale with 0 occurrences must receive a value of `0` in their respective series.

2. **Series Construction**:
   - Build explicit series objects with matching lengths:
     ```json
     [
       { "name": "MITIGADO", "values": [0, 0, 1, 1, 0] },
       { "name": "VULNERABLE", "values": [0, 0, 0, 0, 0] }
     ]
     ```
   - Total series values across all categories must match the sum of table rows exactly.

---

## 3. Data Integrity & Hallucination Prevention

- Never fabricate data rows, findings, or metrics that do not appear in `[FULL_CONTEXT]` or `[SELECTED_CONTEXT]`.
- If a table cell is empty or contains `N/A`, preserve it or treat it as `null` / `0` for numeric calculations.
- Always cross-verify calculated sums with summary narratives.
