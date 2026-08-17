# Architecture decision log

This log is append-only. Do not renumber or silently rewrite accepted decisions. Supersede an earlier decision with a new entry that references it.

## Decision template

### D-XXX — Title

- **Date:** YYYY-MM-DD
- **Status:** Proposed | Accepted | Superseded | Rejected
- **Context:**
- **Decision:**
- **Alternatives considered:**
- **Consequences:**
- **Files/components affected:**
- **Supersedes:** None

---

### D-001 — Container-level inventory granularity

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** The laboratory needs traceability of complete bottles, boxes, and packages but does not measure remaining contents.
- **Decision:** One physical container/package is one inventory unit. No volume, mass, or individual-piece depletion is tracked in v1.
- **Alternatives considered:** Piece-level inventory; volume-level depletion; manually stored quantity balances.
- **Consequences:** Stock is a count of qualifying Container records. Empty contents cannot be inferred and require a recorded event.
- **Files/components affected:** All architecture, formulas, transactions, Scan workflow, Dashboard.
- **Supersedes:** None

### D-002 — Architecture before VBA

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** VBA written against unstable sheet/Table/column names becomes fragile and difficult to verify.
- **Decision:** No VBA may be written or embedded before the macro-free workbook contract is frozen and non-VBA tests pass.
- **Alternatives considered:** Macro-first prototyping; simultaneous workbook and VBA design.
- **Consequences:** Initial delivery is a macro-free `.xlsx`; automation follows only after review.
- **Files/components affected:** Entire repository and delivery sequence.
- **Supersedes:** None

### D-003 — Basic v0.1 may proceed using documented defaults

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** The owner wants a useful basic version without a lengthy questionnaire.
- **Decision:** Use `docs/default-assumptions.md` for non-critical choices, document all assumptions, build a reviewable macro-free v0.1, and stop before VBA.
- **Alternatives considered:** Require all operational details before beginning.
- **Consequences:** Architecture remains editable until explicitly frozen; no avoidable blocking questions.
- **Files/components affected:** Initial task and architecture phase.
- **Supersedes:** None
