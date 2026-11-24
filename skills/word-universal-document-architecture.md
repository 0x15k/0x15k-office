---
name: Word Universal Document Architecture & OOXML Mastery
description: Enforces institutional document structure, dual-split margin alignment, clean table pagination, and native DrawingML OOXML chart data manipulation in Microsoft Word.
hosts: [Word]
alwaysActive: true
keywords: [word, document, architecture, headings, table, chart, ooxml, drawingml, margins, dual-split, executive, report, memo, pagination]
---

## 1. Document Structural Hierarchy

1. **Heading Standard**:
   - Title / Document Name: 20–24pt Bold, centered or left-aligned with a subtle accent bar.
   - Heading 1: 14–16pt Bold, `#1B365D` (Dark Navy) or `#111827` (Charcoal), 12pt space-before, 4pt space-after.
   - Heading 2: 12–13pt Bold, `#2D3748`, 8pt space-before, 2pt space-after.
   - Heading 3: 11pt Bold or Italic, 4pt space-before.
   - Body Paragraphs: 10.5–11pt Regular, line spacing 1.15x (276 twips), 4–6pt space-after. Zero empty paragraph lines (`¶ ¶`).

2. **Dual-Split Margin Alignment (No Broken Tabs)**:
   - For all metadata pairs (e.g. `[Organization] | [Location]` or `[Degree / Role] | [Dates]`):
   - In Markdown, represent as `[Left Entity] | [Right Date/Location]`.
   - The engine automatically transforms these into a borderless 100% width table with cell 1 left-aligned and cell 2 right-aligned.
   - NEVER insert typewriter tabs (`\t` or `\t\t`) or strings of manual spaces (`    `), which break responsive margin reflow in Word.

3. **Institutional Table Standards**:
   - Always set table headers with bold text, subtle fill (`#F3F4F6` or corporate primary tint), and explicit column widths.
   - For multi-page tables, ensure Word repeats headers across page breaks and prevents row splitting:
     - Wrap header row in `<w:tblHeader/>`.
     - Set rows with `<w:cantSplit/>` to eliminate awkward mid-row page splits.
   - Numeric columns must be right-aligned; text columns left-aligned; status/date columns centered.

4. **Executive Callout Boxes**:
   - For key findings, audit observations, warnings, or executive summaries:
   - Construct a single-cell callout table with a 3pt left accent border (`#2563EB` or `#DC2626`), no top/right/bottom borders, 8pt internal padding, and subtle background tint (`#F8FAFC`).

5. **Native DrawingML OOXML Chart Mutation**:
   - When updating numbers, labels, or series data in existing Word charts, use the native `update_chart` operation:
     ```json
     {
       "action": "update_chart",
       "paragraphIndex": 14,
       "chartIndex": 0,
       "categories": ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025"],
       "series": [
         { "name": "Revenue ($M)", "values": [12.4, 15.8, 18.2, 22.5] },
         { "name": "Operating Cost ($M)", "values": [8.1, 9.4, 10.2, 11.0] }
       ]
     }
     ```
   - The chart engine parses the underlying DrawingML XML (`<c:chartSpace>`), updates the series cache (`<c:numCache>` and `<c:strCache>`), purges broken `<c:externalData>` links to detached Excel workbooks, and forces Word to re-render the embedded vector graphics immediately.

6. **Full-Document Context Ingestion**:
   - Always inspect the `[FULL_CONTEXT]` (up to 250,000 characters) and AST outline (up to 500 paragraphs) to maintain cross-section consistency between Executive Summaries, Analysis tables, Chart captions, and Final Recommendations.
