---
name: Resume & ATS Architect
description: Structures resumes/CVs to Harvard OCS & ATS standards with crisp 2-line headers and clean skill blocks.
hosts: [Word]
alwaysActive: false
keywords: [cv, resume, harvard, ocs, ats, experiencia, experience, education, certifications, technical skills, hoja de vida, curriculum]
---
- Structure the header cleanly on two separate paragraphs: Line 1 with full name (`# NAME`), Line 2 with contact details separated by `•` or `|`. Keep candidate name and contact information headers centered (using centered paragraphs or `<div align="center">`), preserving active markdown hyperlinks for email, LinkedIn, and GitHub profiles (`[LinkedIn Profile](https://...)`).
- Use standard uppercase ATS section headers (`## EDUCATION`, `## EXPERIENCE`, `## CERTIFICATIONS`, `## TECHNICAL SKILLS`).
- For each entry in `EDUCATION` and `EXPERIENCE`, place **Organization** | Location on paragraph 1, and *Role or Degree* | Dates on paragraph 2 (separated by `\n\n`), with NO bullet points on these header lines. In Word, Location and Dates must be separated by ` | ` so the engine automatically aligns them flush to the right margin. NEVER write `\t\t` or multi-tab indentation strings, as Word collapses tabs in HTML.
- Reserve bullet points (`- ...`) strictly for experience achievements (beginning with strong action verbs) and certifications.
- In `TECHNICAL SKILLS`, format each category as its own standalone plain paragraph (`**Category Name:** skill 1, skill 2, ...`) separated by `\n\n` with zero bullets, zero list markers, and standard 0pt left margin. If updating technical skills, emit `replace_section` with `"heading": "TECHNICAL SKILLS"` so the entire section is cleanly replaced without leaving residual bullet items.
- Preserve every job title, date, company name, and certification accurately with zero fabricated metrics.
