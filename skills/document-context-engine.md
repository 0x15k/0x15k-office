---
name: Document Context Engine & Positional Intelligence
description: Universal context-awareness, positional hierarchy mapping, conversational deixis disambiguation, and cross-section data synchronization for Microsoft Office across all languages.
hosts: [Word, Excel, PowerPoint]
alwaysActive: true
keywords: [context, hierarchy, position, deixis, disambiguation, cross-reference, consistency, chart, table, caption, section, revert, anterior, previous, language-agnostic]
---

# Universal Document Context Engine & Positional Intelligence

This skill provides the cognitive foundation for interpreting Office document geometry, understanding conversational references, and ensuring cross-section synchronization across documents in any language.

---

## 1. Document Geometry & Positional Hierarchy Mapping

Every professional institutional document follows an explicit spatial hierarchy. Always map the document into its functional zones before issuing edits:

1. **Cover / Title Block**:
   - Institutional identifier, report code, title, subtitle, date, author/reviewer block.
   - Changes here must cascade to headers, footers, and metadata tables throughout the document.

2. **Executive Summary / Management Abstract**:
   - High-level executive synthesis summarizing the findings or business rationale.
   - Typically contains a high-level summary table (e.g. Total Assessed, Status Breakdown).
   - Must mirror the exact aggregates of the technical body with 100% numerical consistency.

3. **Visual Exhibit Triad (Caption → Exhibit → Narrative)**:
   - Visual graphics (charts, architecture diagrams, callout boxes) NEVER exist in isolation.
   - They ALWAYS form a coherent 3-part spatial unit:
     1. **Preceding Descriptive Caption**: Contextualizes what the reader is about to see (e.g. *"El gráfico siguiente muestra la distribución actual de..."*).
     2. **Visual Exhibit / Chart Drawing**: The graphical artifact itself (Pie, Bar, Column, Diagram).
     3. **Succeeding Narrative Analysis**: The interpretive takeaway (e.g. *"El análisis evidencia el nivel de avance en la gestión de riesgos..."*).
   - **Critical Rule**: When replacing, mutating, or alternating a visual exhibit, NEVER touch or delete the preceding caption or the succeeding analysis paragraph unless explicitly asked. The exhibit must remain anchored firmly between them.

4. **Technical Findings / Detailed Evidence**:
   - Comprehensive data rows, technical cards, individual vulnerability/issue write-ups, or itemized line items.
   - This is the source of ground truth for all metrics.

5. **Conclusions & Actionable Next Steps**:
   - Synthesizes findings and defines remediation/next milestones. Must harmonize with both the Executive Summary and the Technical Findings.

---

## 2. Conversational Intent & Deixis Disambiguation

Users frequently issue relative or deictic commands that depend heavily on conversational and spatial context:

1. **Reversion & Alternation Commands**:
   - When the user asks: *"vuelve al gráfico de barras anterior que teníamos"*, *"hazlo circular de nuevo"*, *"reviértelo al formato original"*, or *"vuelve a la tabla"*:
     - **Identify the Active Element**: Locate the visual exhibit or table currently occupying that slot in the document.
     - **Retrieve the Governing Data Series**: Extract the active metrics from the surrounding summary table (e.g., Mitigated: 1 Media, 1 Baja; Vulnerable: 0).
     - **Regenerate the Alternate Representation**: Generate the requested format (e.g., horizontal severity bar chart with categories and legend, or executive donut chart) with the exact same data points.
     - **Surgical In-Place Replacement**: Replace ONLY the active exhibit, maintaining the centered alignment, high-resolution retina scaling, and caption/analysis anchors.

2. **Relative Spatial References**:
   - *"la tabla de arriba"* → The nearest table preceding the current selection or paragraph.
   - *"el párrafo siguiente"* → The paragraph immediately below the referenced object.
   - *"la sección de resultados"* → The section header containing "Resultados", "Hallazgos", or "Findings".

3. **Cross-Referenced Metric Syncing**:
   - If the user modifies a number in a technical finding (e.g. changes a finding from "Vulnerable" to "Mitigado"), the engine MUST automatically propose or execute updates to:
     - The Executive Summary table counts.
     - The Visual Chart data series.
     - The Conclusion paragraph metrics.

---

## 3. Language-Agnostic Reasoning Directive

1. **English Instructions for Precision**:
   - All AI skills, system instructions, AST annotations, and operation schemas are written in English for maximum reasoning fidelity and strict adherence by state-of-the-art models.

2. **Native Document Output**:
   - Regardless of the prompt language or skill instruction language, **the generated document content, table headers, chart titles, labels, legends, and narrative paragraphs MUST strictly match the native language of the active document** (Spanish, English, Portuguese, etc.).
   - Never inject English text, English column titles, or English status words into a Spanish document (e.g. use "MITIGADO" not "MITIGATED", "CRÍTICA" not "CRITICAL", "Distribución de Vulnerabilidades" not "Vulnerability Distribution").
   - Preserve existing corporate capitalization conventions, terminology, and acronyms without artificial translation.

---

## 4. Visual Exhibit Semantics, Metric Selection & Context-Driven Color Protocols

Visual exhibits (charts, graphs, matrices) are not mere cosmetic decorations; they are high-impact cognitive instruments that synthesize business reality for executive decision-makers.

### 4.1 Metric Dimension Selection (Status vs. Severity)
When designing or converting a visual exhibit, the engine must analyze the document's primary business question:

1. **Audit / Recertification / Remediation Context**:
   - **Primary Business Question**: *"Did the remediation succeed? Are there remaining active risks?"*
   - **Governing Executive Metric**: Remediation Status (e.g., `MITIGADO: 100% (2/2)`, `VULNERABLE: 0%`).
   - **Common Pitfall to Avoid**: Slicing an executive pie chart simply by original severity (e.g., `50% Media`, `50% Baja`) without indicating mitigation status misleads executives into believing that half the system remains vulnerable.
   - **Resolution**:
     - *Approach A (Status Breakdown)*: Chart the remediation status as the primary slices (`MITIGADO` vs `VULNERABLE`).
     - *Approach B (Mitigated Severity Breakdown)*: If showing severities in a recertification report, explicitly qualify each category (e.g., `Media (Mitigada): 1`, `Baja (Mitigada): 1`) so status is unambiguous.
     - *Approach C (Stacked Bar Matrix)*: Use stacked bars with categories = Severities and series = `[MITIGADO, VULNERABLE]` showing both dimensions simultaneously.

2. **Initial Assessment / Pentest / Threat Model Context**:
   - **Primary Business Question**: *"What vulnerabilities were found, and what is the severity profile?"*
   - **Governing Executive Metric**: Severity Distribution (`CRÍTICA`, `ALTA`, `MEDIA`, `BAJA`, `INFORMATIVA`).

### 4.2 Context-Driven Color Resolution & Brand Palette Hierarchy
Colors must NEVER be hardcoded in engine logic. They must be resolved dynamically through the following hierarchy:

1. **Tier 1: Active Custom Skills (`[ACTIVE_SKILLS]`)**:
   - Always extract exact corporate hex codes defined in the active skills (e.g., `formato-informe-bch-v2` specifies `MITIGADO: #92D050`, `VULNERABLE: #FF0000`, `CRÍTICO: #501549`, `MEDIO: #FFC000`, `BAJO: #4EA72E`, `INFORMATIVO: #0F9ED5`).
   - Pass these mapped hex values directly into `update_chart` via `colors: { "CategoryOrSeries": "#HEX" }`.

2. **Tier 2: Explicit User Instruction**:
   - If the user requests specific colors (e.g., *"usa azul y naranja"*, *"pon el mitigado en verde oscuro"*), respect user specifications over defaults.

3. **Tier 3: Contextual Query or Universal Professional Palette**:
   - If neither custom skills nor document context provide a color palette, and the visual distinction is critical, **ask the user** what corporate colors they prefer, or apply standard neutral corporate palettes (`#2563EB`, `#10B981`, `#F59E0B`, `#EF4444`) and mention the color scheme in the chat reply.

---

## 5. Expected Operational Steps (Universal Workflow)

When executing any contextual document edit:

- **Step 1: Context Ingestion**: Read the active selection, surrounding paragraphs (±3 siblings), and the section heading to determine exact document coordinates.
- **Step 2: Intent Resolution**: Check if the command is an initial creation, a surgical refinement, or a format alternation/reversion.
- **Step 3: Semantic Metric Audit**: Identify the primary metric dimension (Status vs Severity) and verify that numbers match the governing table.
- **Step 4: Palette Resolution**: Extract color mappings from `[ACTIVE_SKILLS]` or document styles and map them into `op.colors`.
- **Step 5: Surgical Execution**: Issue the appropriate operation (`update_chart`, `replace_section`, `search_and_replace`, `format_selection`) targeting only the affected element.
- **Step 6: Post-Edit Harmonization**: Ensure that formatting (alignment, font, margins, captions) remains pristine and harmonious with the rest of the document.
