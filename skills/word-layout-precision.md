---
name: Word Layout Precision & Micro-Typography
description: Enforces surgical section targeting, distinct paragraph spacing, AST outline grounding, and micro-typographic standards in Microsoft Word.
hosts: [Word]
alwaysActive: true
keywords: [word, layout, precision, typography, margins, spacing, bullets, ast, outline, targeting]
---

## 1. Context Grounding & Surgical Action Targeting

1. **Full AST Outline Navigation**:
   - The engine provides a comprehensive AST outline in `[DOCUMENT_OUTLINE]` spanning up to 500 paragraphs with structural annotations (`(Heading 1)`, `(Heading 2)`, `(Table)`, `(ChartRef)`, `(DualSplit)`).
   - Use the AST outline to precisely locate target paragraphs and sections before issuing edits.
   - NEVER use `replace_selection` when `[SELECTED_CONTEXT]` is empty or does not precisely match the user's intended target. Instead, use `replace_section` with `"heading": "HEADING_TEXT"`, `search_and_replace`, or `replace_full_document`.

2. **Full-Document Ingestion (up to 250,000 Chars)**:
   - When asked to analyze, standardize, or summarize multi-page reports (up to 60–80 pages), inspect the complete `[FULL_CONTEXT]`.
   - Ensure consistency across all sections: terminology, capitalization, citation format, numbering, and cross-references.

## 2. Micro-Typography & Vertical Rhythm

1. **Paragraph Vertical Rhythm**:
   - Distinct paragraphs MUST be separated by double newlines (`\n\n`) in Markdown.
   - In Word, spacing between paragraphs is controlled by `spaceBefore` and `spaceAfter` properties (typically 6pt after, 0pt before).
   - NEVER insert empty paragraphs (`¶ ¶` or `\n\n\n\n`) to simulate vertical spacing; empty paragraphs create unpredictable page breaks and uneven margins.

2. **Typographic Proportions**:
   - Line Spacing: Standard professional line spacing is **1.15x multiple** (276 twips) or **1.2x multiple** for long-form prose.
   - Font Consistency: Maintain the font family detected in `[DOCUMENT_STYLE_STATE]` (e.g. Aptos, Calibri, Arial, Times New Roman). Never mix incompatible fonts across adjacent paragraphs.
   - Casing: Maintain uniform title casing on all parallel headings.

3. **List Attributes vs Text Characters**:
   - In Microsoft Word, list bullets (`•`, `-`, numbered lists) are native paragraph-level formatting attributes, NOT character glyphs.
   - Never generate multi-line search targets containing fake virtual tab stops or literal bullet glyphs (`•\t`).
   - When converting bullet lists to clean paragraphs or vice versa, use `replace_section` or `replace_full_document` so the engine binds the native Word list attributes cleanly.

4. **Preserving Intentional Formatting**:
   - Preserve all existing hyperlinks (`[Text](url)`).
   - Preserve centered alignment on headers, title blocks, and signature blocks.
   - Maintain table column proportioning and cell borders.
