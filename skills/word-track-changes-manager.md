---
name: Word Track Changes & Redline Manager
description: Enterprise management of Word Track Changes (revisions), markup views (All Markup, Simple Markup), redline audits, non-destructive OOXML replacements, and revision acceptance/rejection workflows.
hosts: [Word]
alwaysActive: true
keywords: [track changes, control de cambios, revisions, redline, all markup, simple markup, accept changes, reject changes, balloon, deleted, inserted, ooxml, reviewing]
---

# Word Track Changes & Redline Manager Skill

Adapted from the GenOffice revision engine (`accept_changes`, `reject_changes`, `docs compare`) and Microsoft Word OpenXML revision architecture, this skill governs how the AI operates when documents have active Track Changes.

---

## 1. Word Track Changes (Revision Mode) Architecture

In Microsoft Word, when Track Changes (`Reviewing` mode) is enabled:

1. **Non-Destructive Delta Storage**:
   - Word does NOT overwrite or erase deleted content on disk.
   - Deleted content is wrapped in `<w:del>` or `<w:delText>` and attributed to the current author with an ISO timestamp.
   - Inserted content is wrapped in `<w:ins>` with author metadata.
   - In paragraphs with paragraph-level formatting changes, a `<w:pPrChange>` element is recorded.

2. **Visual Markup Modes in Word UI**:
   - **All Markup**: Word renders all insertions inline (underline / colored) and all deletions in right-margin balloons or strikethrough.
   - **Simple Markup**: Word shows a clean preview with a red vertical line in the left margin indicating modified paragraphs.
   - **No Markup**: Word displays the document as it would appear if all changes were accepted.
   - **Original**: Word displays the document prior to any changes.

---

## 2. Handling Visual Exhibits & DrawingML Under Track Changes

When modifying or replacing visual exhibits (such as DrawingML charts, diagrams, or tables) while Track Changes is enabled:

1. **Why Deleted Exhibits Persist on Screen**:
   - When a paragraph containing a DrawingML chart is replaced via `insertOoxml(..., replace)`, Word's Track Changes engine marks the previous paragraph as a tracked deletion.
   - If the user is viewing **All Markup**, Word displays a balloon on the right margin saying:
     `[Author] [Timestamp] - Deleted: [Chart Graphic]`
   - This balloon is NOT an error. It is Word's native audit trail documenting the previous state.

2. **Avoiding Duplicate Visual Artifacts**:
   - If the user asks for a *"clean view"* or says *"quita los cambios"* / *"acepta los cambios"*:
     - Explain clearly that the document is in **Reviewing (Track Changes)** mode.
     - Instruct the user to switch the display dropdown in the **Review** tab from **All Markup** to **Simple Markup** or **No Markup**, or click **Accept → Accept All Changes** to permanently bake in the modifications.

3. **Replacing Charts Cleanly**:
   - When issuing `update_chart`, target the exact paragraph holding the exhibit.
   - Never insert a second paragraph before or after unless creating a new comparative exhibit.
   - The exhibit remains strictly anchored between its preceding caption and succeeding narrative analysis.

---

## 3. Redline & Review Workflows

When the user asks to review or redline a document:
- Preserve existing comment threads (`<w:commentRangeStart>`, `<w:commentRangeEnd>`).
- When proposing corrections to legal, technical, or financial contracts, maintain tracked changes so legal counsel and stakeholders can inspect every insertion and deletion individually.
- For executive presentations and final sign-offs, advise on consolidating revisions to maintain visual clarity.
