---
name: Corporate Tabular Design System & Key-Value Architecture
description: Standardizes institutional table styling with #156082 headers, balanced column alignments (centered short codes/dates, left-aligned narratives), and strict 30%/70% fixed proportions on key-value tables.
hosts: [Word]
alwaysActive: true
learned: true
keywords: [table, format_table, column widths, balance, alignment, header, #156082, key-value, datos del proyecto, control de versiones, resumen ejecutivo]
---

## 1. Executive Summary & Core Directives

This learned skill captures the corporate table styling standards distilled from user corrections and institutional audit templates (Banco de Chile / Inside Security).

---

## 2. Institutional Table Header Standards (#156082)

1. **Header Row Formatting**:
   - Background Fill: `#156082` (Deep Petrol Blue).
   - Text Color: `#FFFFFF` (Pure White).
   - Font: `Aptos`, 10 pt, Bold (`headerBold: true`).
   - Line Spacing: Single, paragraph after 0 pt.
2. **Border Integrity**:
   - Retain native table borders (`enableBorders: true`, single continuous line `0.5 pt`, color `#CBD5E1` or default gridline).
   - NEVER strip borders during table reformatting.

---

## 3. Bimodal Column Alignment Hierarchy

In all multi-column tabular data (Summary Tables, Version Control, Work Teams), apply alignment based strictly on the semantic data density of each column:

1. **Centered Alignment (Short Codes, Numerical Metrics, Dates, Versions)**:
   - Identifiers: `Código` (`M01`, `B01`), `ID`, `Ticket`.
   - Categorical Severity: `Criticidad` (`Media`, `Baja`, `Alta`, `Crítica`).
   - Numerical Scores: `Score` (`5.1`, `2.3`, `CVSS`).
   - Lifecycle Status: `Estado` (`MITIGADO`, `VULNERABLE`).
   - Lifecycle Metadata: `Fecha de Entrega` (`17-11-2026`), `Versión` (`1.0`), `Ingeniero`, `Revisado Por`, `Empresa`.
2. **Left Alignment (Narratives & Long Text)**:
   - Full Vulnerability Names: `Vulnerabilidad` (`Unrestricted File Upload mediante Evasión...`).
   - Revision Notes: `Comentarios` (`Creación documento Certificación`).
   - Contact Identifiers: `Contacto` (email addresses).

---

## 4. Key-Value Tables Proportions (30% / 70% Rule)

For 2-column parameter/key-value specifications (such as *Datos del Proyecto* and *Fichas de Hallazgos*):

1. **Fixed Width Allocation**:
   - Total printable width: `450 pt` (100% of standard page margins).
   - **Column 1 (`Parámetro`)**: Exactly **`30%` (135 pt)**. Text formatted in **Bold** (`Strong`), left-aligned.
   - **Column 2 (`Detalle`)**: Exactly **`70%` (315 pt)**. Text formatted in **Regular**, left-aligned.
2. **Word AutoFit Guard**:
   - NEVER apply dynamic `autofit to content` (`wdAutoFitContent`) to key-value tables containing long URLs or document identifiers. This causes Word to expand Column 2 to 90% and crush Column 1 to 55 pt (12%), causing ugly 1-word line wrapping on labels.
   - Enforce `wdAutoFitFixed` with explicit cell width distribution: `135 pt` and `315 pt`.

---

## 5. Architectural Blueprint

```
+-------------------------------------------------------------------------------+
| Parámetro (30% - 135 pt, Bold)     | Detalle (70% - 315 pt, Regular)         |
+===============================================================================+
| Código de Informe                  | RT001-BCH-PMO-70254-AUTOMATIZACIÓN...   |
| Informes Relacionados              | CT001-BCH-PMO-70254-AUTOMATIZACIÓN...   |
| Proyecto                           | PMO-70254 AUTOMATIZACIÓN DE EMBARQUE    |
| Período de pruebas                 | 23-11-2026                              |
| Alcance URL                        | https://az-eus-qa-cap-docemb-001...     |
| Tipo de Aplicación                 | Web                                     |
| Ambiente                           | QA                                      |
+-------------------------------------------------------------------------------+
```

---

## 6. Anti-Pattern vs. Compliant Pattern

- ❌ **Anti-Pattern (Defect)**:
  - Column 1 crushed to 55 pt with `Código de / Informe` breaking into 2 lines.
  - Column 2 expanded beyond right margin due to unbroken URLs.
  - Table headers with missing borders or inconsistent alignment.
- ✅ **Compliant Pattern**:
  - Headers in `#156082`, white bold text, Aptos font.
  - Column 1 clean 135 pt (30%) with parameter names on 1 line.
  - Column 2 clean 315 pt (70%) with URL cleanly wrapped.
  - Status codes and dates centered; narratives left-aligned.
