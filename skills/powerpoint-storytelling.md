---
name: PowerPoint Executive Storytelling & Slide Architecture
description: Builds executive-ready presentation decks with action headlines, MECE narrative structure, modular visual cards, KPI callouts, and speaker notes.
hosts: [PowerPoint]
alwaysActive: true
keywords: [powerpoint, ppt, slide, deck, presentation, pitch, storytelling, pyramid principle, mece, executive, speaker notes]
---

## 1. Executive Narrative Architecture (Pyramid Principle)

1. **Action Titles Over Descriptive Labels**:
   - ❌ Weak: "Financial Overview" or "Market Trends".
   - ✅ Strong: "Revenue Grew 24% YoY Driven by Enterprise Cloud Migration".
   - Every slide title must formulate the core conclusion or takeaway in 8–14 words.
   - Subtitle / Tracker: 1-line context or category tag directly above or below the title (e.g. `MARKET ANALYSIS | 2026 FORECAST`).

2. **MECE Structural Scaffolding**:
   - Structure deck arguments to be Mutually Exclusive and Collectively Exhaustive (MECE).
   - Standard slide narrative rhythm:
     - **Slide 1**: Executive Context & Burning Platform (The "Why Now").
     - **Slide 2**: Strategic Findings / Diagnosis (The Problem Statement).
     - **Slide 3**: Proposed Solution Architecture / Framework (The "What").
     - **Slide 4**: Operational Roadmap & Phased Execution (The "How").
     - **Slide 5**: Financial Impact, ROI & Risk Mitigation (The "Value").

## 2. Visual Layout Coordinates (16:9 Widescreen Standard)

PowerPoint 16:9 widescreen dimensions are **960 x 540 points** (or 13.33 x 7.5 inches):
- **Slide Title**: `left: 50, top: 40, width: 860, height: 45` (Font: 22–26pt Bold).
- **Subtitle / Action Tracker**: `left: 50, top: 88, width: 860, height: 25` (Font: 12–14pt, `#6B7280`).
- **Main Content Zone**: `top: 130 to 470` (Height: 340pt).

### Modular Visual Layout Patterns:

- **3-Column Card Layout** (Comparison / Pillars):
  - Card 1: `left: 50, width: 265, top: 130, height: 330`
  - Card 2: `left: 345, width: 265, top: 130, height: 330`
  - Card 3: `left: 640, width: 265, top: 130, height: 330`
  - Background fill: `#F8FAFC` or subtle border `#E2E8F0` with 8pt internal padding.

- **KPI Metric Callout Cards**:
  - Top number: 36–44pt Bold (`#1B365D` or `#2563EB`).
  - Metric label: 11pt Uppercase Charcoal (`#4B5563`).
  - Delta / Benchmark: 10pt (`▲ +18% vs prior year` in green `#16A34A` or `▼ -4%` in red `#DC2626`).

- **Key Takeaways Callout Bar**:
  - Horizontal container at bottom: `left: 50, top: 460, width: 860, height: 50`
  - Fill: `#EFF6FF` (Ice Blue), Left Accent Border 4pt `#2563EB`.

## 3. Executive Delivery & Speaker Notes

- When generating slides via `create_presentation` or `add_slide`, always provide detailed `speakerNotes`:
  - 1st sentence: The governing thesis of the slide.
  - 2nd–3rd sentences: Key supporting data points and context not displayed on the visual card.
  - Anticipated executive questions and answers (FAQ prep).
- Bullet discipline: Maximum 4 bullets per text container, each under 15 words. Bold the first 2–4 words of each bullet for visual scanning.
