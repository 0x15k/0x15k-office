---
name: Universal Word Document Architecture & Dual Alignment
description: Enforces consistent typographic hierarchy, heading styles, and dual-split right-alignment for entities, locations, and dates across all sections in Microsoft Word documents.
hosts: [Word]
alwaysActive: true
---

## Rules
- MUST format all major document section headings (e.g., EDUCATION, EXPERIENCE, PROJECTS, SUMMARY, HONORS & AWARDS, CERTIFICATIONS) consistently as Heading 2 with bold styling, uppercase or title casing, 6pt space-before, and no residual tab characters.
- MUST use a consistent dual-alignment structure across all parallel entries (such as borderless dual-split tables or edge-to-edge layout) to anchor secondary metadata (locations, dates, credentials, status) flush to the right margin.
- MUST maintain uniform metadata granularity across all sections (e.g., consistently use '[City, Country/Region]' or '[Location / Remote]', without mixing disparate geographic or temporal formats within the same document).
- MUST apply uniform typographic hierarchy across peer entries (e.g., bold for primary organization/institution names, italics for roles/degrees/sub-headers, regular text for descriptions).
- MUST NEVER use raw inline typewriter tab stops ('\t') or plain pipe runs that collapse or cause asymmetrical left-alignment in Microsoft Word.
- MUST enforce clean list and bullet hygiene:
  * In dedicated list sections (such as HONORS & AWARDS, AWARDS, CERTIFICATIONS, ACHIEVEMENTS), EVERY single item MUST be formatted as an active bullet item ('- Item'). NEVER leave any award, competition ranking, or certification as unbulleted plain text.
  * Bullet items must NEVER be double-prefixed with arrows, dashes, or manual symbols (write '- 1st Place...', NOT '- -> 1st Place...' and NOT '• -> 1st Place...').
  * Ghost bullets ('• ¶') containing no substantive text must be completely eliminated.
- MUST preserve centered alignment on author/candidate names and contact headers, keeping active clickable hyperlinks ('[Display Text](url)') intact.
- MUST control vertical document rhythm using paragraph properties ('spaceBefore', 'spaceAfter') and line spacing multipliers (e.g., 1.15x multiple), never through repeated blank paragraphs ('¶ ¶').

## Examples

### ❌ Anti-pattern (Before / Rejected)
```
[Primary Entity / Organization] | [Location]
[Role / Degree Title]	[Date Range]
• -> [Achievement item with duplicate arrow]
• 
```

### ✅ Correct Pattern (After / Compliant)
```
| **[Primary Entity / Organization]** | [Location / Region] |
| :--- | ---: |
| *[Role / Degree Title]* | *[Date Range]* |

- [Clean achievement statement with active action verb]
```
