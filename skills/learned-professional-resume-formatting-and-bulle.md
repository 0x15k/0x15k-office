---
name: Professional Document Layout and Bullet List Standardization
description: Standardizes structural hierarchy, dual-aligned metadata headers, clean bullet lists, and typographic spacing across professional documents in Microsoft Word.
hosts: [Word]
alwaysActive: true
---

## Rules
- MUST replace manual tab characters (`\t`) and multiple blank lines with structured two-column borderless tables or native tab stops for dual-aligned metadata (e.g., [Entity / Institution] left-aligned and [Location / Region] right-aligned; [Role / Degree] left-aligned and [Date Range] right-aligned).
- MUST use native paragraph formatting properties (spaceBefore and spaceAfter) to maintain vertical rhythm instead of inserting empty paragraph breaks.
- MUST ensure list items under credential, honor, or experience sections are genuine, uniform list items with standard bullet glyphs, rather than plain text prefixed with manual hyphens, asterisks, or duplicate arrow glyphs.
- MUST maintain uniform typographic scale, font family, line spacing, and capitalization across all section headers and sub-elements.
- MUST NEVER hardcode personal identities, locations, or institution names, maintaining geographic and domain neutrality across all templates and structural elements.
- MUST NEVER mix unbulleted run-on blocks and pseudo-bulleted text within list-driven sections such as honors, awards, or certifications.

## Examples

### ❌ Anti-pattern (Before / Rejected)
```
## EDUCATION
[Institution Name]
[Location]

[Degree Title]
[Date Range]

## CERTIFICATIONS
-> [Certification Name 1]
*   [Certification Name 2]
[Certification Name 3] - [Issuing Body]
```

### ✅ Correct Pattern (After / Compliant)
```
## EDUCATION
| [Institution Name] | [Location / Region] |
| [Degree / Qualification Title] | [Date Range] |
- [Academic Distinction or Relevant Coursework]

## CERTIFICATIONS
- [Certification Title 1] – [Issuing Organization], [Year]
- [Certification Title 2] – [Issuing Organization], [Year]
- [Certification Title 3] – [Issuing Organization], [Year]
```
