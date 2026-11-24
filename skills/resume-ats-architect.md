---
name: Resume & ATS Architect (Harvard OCS Standards)
description: Formats executive and technical resumes to Harvard OCS & ATS standards with crisp 2-line dual-split headers, XYZ achievement bullets, and zero parsing traps.
hosts: [Word]
alwaysActive: false
keywords: [cv, resume, harvard, ocs, ats, curriculum, hoja de vida, experience, education, certifications, technical skills, bullet]
---

## 1. Structural Architecture & ATS Parsing Compliance

1. **Header Layout**:
   - Line 1: Candidate Name (`# CANDIDATE FULL NAME` or 18–22pt Bold, centered).
   - Line 2: Contact row centered (`City, Country • phone • email • [LinkedIn Profile](url) • [GitHub Profile](url)`).
   - Preserve clickable hyperlinks for email, LinkedIn, and portfolios. Never convert links to plain text.

2. **Standard Section Hierarchy**:
   - Use standard uppercase ATS headings:
     - `## PROFESSIONAL SUMMARY` (Optional, 3–4 punchy lines)
     - `## EXPERIENCE`
     - `## EDUCATION`
     - `## CERTIFICATIONS & HONORS`
     - `## TECHNICAL SKILLS`
   - Zero multi-column tables for the main body layout (ATS parsers read multi-column layouts horizontally across column boundaries, jumbling career chronologies).

3. **Dual-Split Alignment for Positions & Degrees**:
   - For every entry under `EXPERIENCE` and `EDUCATION`, emit two crisp paragraphs separated by `\n\n`:
     - Paragraph 1: `**Company or University Name** | Location`
     - Paragraph 2: `*Job Title or Degree Program* | Month Year – Month Year`
   - In Word, the engine automatically converts `[Left Entity] | [Right Date/Location]` into a borderless 100% width table anchored flush to the margins.
   - NEVER insert manual typewriter tabs (`\t` or `\t\t`), which cause unpredictable indentation shifts when opened across different Word versions.

## 2. Bullet Quality: The XYZ Achievement Framework

- Every bullet under Experience must follow the Google / Harvard XYZ framework:
  `"Accomplished [X], as measured by [Y], by doing [Z]"`
- Start every bullet with a strong, past-tense action verb (e.g., *Architected, Spearheaded, Engineered, Orchestrated, Optimized, Automated, Accelerated*). Avoid passive phrasing ("Responsible for", "Helped with").
- Include tangible metrics where supported by the source text (percentages, latency reductions, dollar amounts, scale).
- Never fabricate numbers or inflate accomplishments (abide strictly by `factual-integrity.md`).

## 3. Technical Skills Block Organization

- Format each skill group as a clean paragraph separated by `\n\n`:
  `**Languages & Frameworks:** Python, TypeScript, Go, React, Node.js`
  `**Cloud & Infrastructure:** AWS (ECS, Lambda, S3), Docker, Kubernetes, Terraform`
  `**Security & Compliance:** OWASP Top 10, Threat Modeling, CIS Benchmarks, Semgrep`
- Zero bullet points on skills lines; clean inline lists separated by commas.
