---
name: DOCX Manipulation & OOXML Architecture
description: Programmatic Word document manipulation, OpenXML package structure, paragraph formatting, table styling, section breaks, and native DrawingML embedding.
hosts: [Word]
alwaysActive: true
keywords: [docx, ooxml, openxml, drawingml, table, paragraph, section, header, footer, style, formatting, flat opc]
---

# DOCX Manipulation & OOXML Architecture Skill

Adapted from the Claude Office Skills `docx-manipulation` module and GenOffice's `docx-engine`, this skill defines the standards for manipulating Microsoft Word documents via Word.js and OpenXML (OOXML).

---

## 1. Document Anatomy & Hierarchy

A Word document (`.docx`) is an Open Packaging Conventions (OPC) ZIP archive containing:
```
word/
├── document.xml          # Main document story: paragraphs (w:p), tables (w:tbl), drawings (w:drawing)
├── styles.xml            # Named paragraph and character styles (Heading 1..9, Normal, etc.)
├── numbering.xml         # Multi-level bullet and numbered list definitions (w:num, w:abstractNum)
├── charts/               # Native DrawingML chart parts (chart1.xml, chart2.xml, ...)
├── media/                # Embedded raster images (image1.png, ...)
├── theme/theme1.xml      # Color scheme and major/minor font definitions
└── _rels/document.xml.rels # Relationship mappings linking rId to parts
```

---

## 2. Flat OPC XML (`pkg:package`) Injection

Word.js allows injecting complete OOXML packages into the active document using `Range.insertOoxml(flatOpc, Word.InsertLocation)`:

1. **Self-Contained Packages**:
   - The XML must declare the `mso-application progid="Word.Document"` processing instruction.
   - It packages `word/document.xml`, `word/_rels/document.xml.rels`, and any referenced `word/charts/chart1.xml` into a single XML envelope under `xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"`.

2. **Native Ribbon Integration**:
   - When a native DrawingML chart (`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`) is inserted via Flat OPC, Microsoft Word activates the official **Chart Design** and **Format** tabs in the ribbon.
   - This allows the user to click on the chart in Word, edit data in Excel, change chart palettes, and format plot areas natively.

---

## 3. Professional Table Formatting Rules

When creating or mutating tables in Word:
- **Header Rows**: Set `<w:tblHeader/>` on row 0 so table headers repeat automatically across page splits.
- **Cell Margins**: Maintain 6pt top/bottom and 8pt left/right padding (`w:tcMar`) for readability.
- **Alternating Shading**: Use subtle zebra banding (`#F8FAFC` vs `#FFFFFF`) for tables exceeding 5 rows.
- **Alignment**:
  - Codes, IDs, Dates, and Status badges: Centered.
  - Descriptions, Names, Findings: Left-aligned.
  - Numbers, Amounts, Scores, Percentages: Right-aligned.

---

## 4. Typography & Line Spacing Preservation

- Always inherit the predominant document typography specified in `[DOCUMENT_STYLE_STATE]`.
- Avoid mixing arbitrary fonts (e.g. Times New Roman inside an Aptos/Calibri document).
- Maintain multiple line spacing (1.15x or 1.2x) and paragraph spacing after (`2pt` to `4pt`) to ensure consistent visual vertical rhythm across pages.
