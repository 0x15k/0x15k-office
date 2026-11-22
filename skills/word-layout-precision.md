---
name: Word Layout Precision
description: Enforces safe section targeting, distinct paragraph breaks, and clean Markdown formatting in Word.
hosts: [Word]
alwaysActive: true
---
- NEVER use `replace_selection` when `[SELECTED_CONTEXT]` is empty or does not match the section the user wants to modify. Instead, use `replace_section` (specifying `"heading"`), `search_and_replace`, or `replace_full_document`.
- In Microsoft Word, list bullets (`•`, `-`, numbers) are paragraph-level list attributes, NOT text characters. NEVER generate multi-line `search_and_replace` queries spanning multiple paragraphs or containing virtual bullets (such as `•\t`). To update or clean formatting across an entire section, use `replace_section` specifying `"heading": "SECTION_NAME"` or `replace_full_document`.
- To ensure Word places consecutive items on separate lines (for example, 2-line Education/Experience headers or `**Category:** items` skill lines), ALWAYS separate distinct paragraphs with blank lines (`\n\n`). Never put multiple `Category: ...` entries on the same paragraph without a blank line between them.
- When restructuring or reformatting an entire document, prefer a single `replace_full_document` paired with `standardize_document` rather than fragmented piecemeal edits.
- When applying typography and line spacing, specify multiple line spacing as a multiplier (e.g., `lineSpacing: 1.15` for 1.15x multiple, `1.5` for 1.5x) in `standardize_document` or `replace_full_document`; never supply arbitrary point heights for standard line spacing multipliers.
- Preserve all existing hyperlinks in the document. Whenever reading, editing, or rewriting text with hyperlinks, format them as `[Display Text](url)` in Markdown so they are inserted as clickable native Word hyperlinks. Never convert hyperlinks to plain text.
- Maintain existing intentional alignment on headers (such as centered candidate names and contact information lines in CVs). Do not reset centered headers to left alignment during document standardization.
- Always inspect [DOCUMENT_STYLE_STATE] and inherit the predominant font, size, and multiple line spacing across all subsequent edits and turns. Never revert to default fonts or arbitrary line heights.
