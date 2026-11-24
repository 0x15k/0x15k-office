---
name: Contract Redline Review & Legal Risk Auditing
description: Performs rigorous contract clause audits, risk scoring (High/Medium/Low), surgical redlines, and native Word review comments for commercial and legal agreements.
hosts: [Word]
alwaysActive: false
keywords: [contract, contrato, clause, cláusula, legal, nda, sla, redline, liability, indemnification, warranty, termination, agreement, terms, review]
---

## 1. Legal Risk Triage Matrix

When auditing contracts, agreements, or master service agreements (MSAs), categorize findings systematically:

- **🔴 HIGH RISK (Deal Breakers / Material Exposure)**:
  - *Uncapped Liability*: Absence of an aggregate liability cap or exclusions that eviscerate the cap.
  - *Broad Indemnification*: One-sided indemnity covering indirect/consequential damages or uncapped IP claims without defense control.
  - *Automatic IP Assignment*: Clauses that inadvertently assign pre-existing background IP, tooling, or know-how to the counterparty.
  - *Immediate Termination for Convenience*: Permitting termination without cure periods or without payment for committed work.

- **🟡 MEDIUM RISK (Commercial Imbalance / Operational Friction)**:
  - *Warranty Scope & Remedies*: Exclusive remedy disclaimers that leave the party with no recourse for chronic SLA failures.
  - *Unilateral Audit Rights*: Excessive audit scope without reasonable prior written notice (minimum 15 business days) or during regular business hours.
  - *Restrictive Non-Solicitation / Non-Compete*: Overly broad geographic or temporal restrictions.
  - *Governing Law & Forum*: Selection of distant or unfavorable jurisdictions without neutral arbitration options.

- **🟢 LOW RISK (Administrative & Clarity Improvements)**:
  - Vague notice delivery definitions (e.g. failure to specify email notice effectiveness).
  - Ambiguous payment milestone conditions or missing net-30 day clarification.

## 2. Redline Execution Standards

1. **Precision Surgical Replacements**:
   - Always use `search_and_replace` with verbatim target strings extracted directly from `[FULL_CONTEXT]` or `[SELECTED_CONTEXT]`.
   - Never replace entire multi-page sections when only a sentence or qualifier needs modification.
   - Maintain the existing contract style, capitalization, and numbering scheme (e.g. `Section 12.2(b)`).

2. **Native Word Review Comments (`add_comment`)**:
   - For every flagged clause, attach an informative review comment using `add_comment`:
     - **Anchor**: Exact phrase in the document (e.g. `"in no event shall Provider's total aggregate liability exceed"`).
     - **Comment Content**:
       `[RISK: HIGH] Uncapped exposure. Recommendation: Cap aggregate liability at 12 months of fees paid under the applicable SOW, excluding gross negligence and willful misconduct.`

3. **Balanced Fallback Drafting**:
   - Always draft commercially reasonable fallback positions rather than aggressive unilateral language that would be immediately rejected in counter-negotiations.
   - Example (Liability Cap):
     `"Except for damages arising from a party's breach of confidentiality, gross negligence, or willful misconduct, each party's total aggregate liability arising out of or related to this Agreement shall be limited to the total fees paid or payable by Client in the twelve (12) months preceding the claim."`
