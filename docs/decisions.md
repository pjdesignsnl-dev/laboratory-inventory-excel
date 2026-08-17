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

### D-004 — Desktop Excel verification deferred to owner environment

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** The build machine has no Microsoft Office/LibreOffice/WPS installation, so desktop Excel cannot run here without changing system configuration.
- **Decision:** Non-VBA testing in this phase uses (1) openpyxl structural inspection, (2) independent formula evaluation with the `formulas` library, and (3) a LibreOffice headless open/recalc smoke test where obtainable. Desktop Excel open/recalc/visual verification is a defined owner-side step before the contract freeze. No claim of desktop Excel testing is made.
- **Alternatives considered:** Installing Office (changes system config; not authorized); skipping verification (violates evidence rules).
- **Consequences:** Evidence labels tests accurately as structural/independent/LibreOffice-smoke; owner performs the desktop gate.
- **Files/components affected:** `tests/non-vba-results.md`, `docs/requirements-analysis.md`, `evidence/`.
- **Supersedes:** None

### D-005 — Status set is 7 values; `Opened` is a flag, not a status

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §11 lists candidate statuses including both `Opened/In Use` and `Available`, and requires distinguishing "previously opened" from "currently taken/in use".
- **Decision:** Use exactly `Available`, `InUse`, `Reserved`, `Expired`, `Damaged`, `Disposed`, `Missing`. Track opening via `OpenedDate`; a returned opened container may be `Available` again while retaining opened history. `Available` is the only stock-qualifying status.
- **Alternatives considered:** An `Opened` status (would block return-to-available); a combined `Taken/In Use`; adding `Empty`.
- **Consequences:** Stock formula is `Status = "Available"`; `OpenedDate` is an independent flag; `Empty` is not modeled (cannot be inferred).
- **Files/components affected:** `tblContainers`, `tblStatusList`, transition matrix, stock formulas.
- **Supersedes:** None

### D-006 — Only `Available` qualifies as stock

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §12 defines stock as containers that "qualify as available"; the phrase is ambiguous about `Reserved`/`InUse`.
- **Decision:** `Available` is the single stock-qualifying status. `Reserved` and `InUse` do not count as stock.
- **Alternatives considered:** Counting `Available` + `Reserved`; counting `Available` + `InUse`.
- **Consequences:** Reorder/stock figures are conservative and match the paper process (only shelf-available units).
- **Files/components affected:** Stock, reorder, dashboard formulas.
- **Supersedes:** None

### D-007 — Nine transaction types; corrections via `Adjustment`

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §10 lists nine transaction concepts and requires compensating/correction behavior.
- **Decision:** Implement `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment`. Corrections are `Adjustment` transactions (append-only); history is never overwritten (invariant 7).
- **Alternatives considered:** Merging expiry/damage/missing into `Dispose`; a separate `Correction` type without state snapshots.
- **Consequences:** Full forensic reconstruction from `tblTransactions`; reversal paths explicit in the matrix.
- **Files/components affected:** `tblTransactions`, `tblTransactionTypeList`, transition matrix.
- **Supersedes:** None

### D-008 — Barcodes are 7-digit zero-padded text; IDs are `C######`/`P######`/`T########`/`S######`/`LOC####`

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Default assumptions 7–9 require text barcodes and immutable IDs distinct from the barcode.
- **Decision:** Barcode `\d{7}` stored as text (leading zeroes preserved, no scientific notation). ContainerID `C######`, ProductID `P######`, TransactionID `T########`, SupplierID `S######`, LocationID `LOC####` — all text, immutable once created.
- **Alternatives considered:** Numeric barcodes; single shared ID; UUIDs.
- **Consequences:** Text-safe formulas; zero-padding preserved; scanner Enter suffix maps digits directly.
- **Files/components affected:** ID columns, validation, Scan/Receiving next-ID formulas.
- **Supersedes:** None

### D-009 — Stock/reorder/expiry are formula-derived, never stored

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §12 and default assumption 10: a separately editable stock number is undesirable.
- **Decision:** Available stock, reorder classification, and expiry bands are computed with COUNTIFS/COUNTIF/IF against `tblContainers`/`tblSettings`. No stock/balance column exists anywhere.
- **Alternatives considered:** Stored balances; event reconstruction only.
- **Consequences:** Integrity preserved; formulas are the single source; dashboard reconciliation is testable.
- **Files/components affected:** Dashboard, Scan, Products, Containers, `tblSettings`.
- **Supersedes:** None

### D-010 — Batch/lot modeled as a Container attribute; no batch table in v1

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §7/§8 mention batch-level data; Req §13 requires batch/lot per chemical. A container is a lot-level physical unit in this design.
- **Decision:** Lot/batch (`BatchLotNumber`), expiry, retest live on `tblContainers`. No separate batch table in v1.
- **Alternatives considered:** A `tblBatches` table keyed by (ProductID, Lot).
- **Consequences:** Simpler v1; batch-level analytics deferred; documented as a future option.
- **Files/components affected:** `tblContainers`, Receiving, Scan.
- **Supersedes:** None

### D-011 — Nine worksheet topology fixed

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §19 suggests a starting sheet set.
- **Decision:** Sheets: `Dashboard`, `Scan`, `Receiving`, `Products`, `Containers`, `Transactions`, `Suppliers`, `Locations`, `Settings`. Nine Tables; staging Tables on Scan/Receiving are non-authoritative UI-only.
- **Alternatives considered:** Merging Suppliers/Locations into Settings; extra "Batch" sheet.
- **Consequences:** Exact sheet/Table names are contract-bound before VBA.
- **Files/components affected:** Entire workbook.
- **Supersedes:** None

### D-012 — Excel 2021/2024-compatible formulas only

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Target environment is 2021/2024 desktop plus Microsoft 365 (default assumption 15).
- **Decision:** All live formulas use COUNTIFS/COUNTIF/IF/AND/OR/TODAY/EDATE/TEXT/MAX/SUMPRODUCT. No `FILTER`/`XLOOKUP`/`SORT`/`LET`/dynamic arrays in cells; 365-only alternatives are documented as side notes.
- **Alternatives considered:** Using 365-only functions for brevity.
- **Consequences:** Maximum compatibility; slightly longer formulas.
- **Files/components affected:** All formula cells.
- **Supersedes:** None

### D-013 — Deployment: single-writer master `.xlsm` on a network share (Option A), read-only copy later

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §20 and default assumption 4/5: one authoritative editor, multiple readers.
- **Decision:** Option A primary: one master workbook on a private network share, written only by the dedicated laboratory PC. Publish a read-only copy (Option B pattern) after v1.0.0 stabilizes. Option C (SharePoint/OneDrive) and Option D (Drive sync) are rejected for the writer path and documented with trade-offs.
- **Alternatives considered:** A, B, C, D (full comparison in `docs/architecture.md` §14).
- **Consequences:** File locking and backup rules defined; read-only reporting artifact deferred.
- **Files/components affected:** Deployment docs, `docs/architecture.md`.
- **Supersedes:** None

### D-014 — Staging Tables on Scan/Receiving are non-authoritative UI-only

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** The Scan sheet must display lookup results without VBA; a formula-driven staging area is the cleanest macro-free mechanism.
- **Decision:** `tblScanResults` and `tblReceiveStaging` hold formula outputs only. They are never sources of truth and are excluded from stock/history logic.
- **Alternatives considered:** Direct cell references without a Table.
- **Consequences:** Consistent structured references; staging rows re-evaluate on input change.
- **Files/components affected:** Scan, Receiving, contract YAML.
- **Supersedes:** None

### D-015 — Sheet protection with empty password; deterrent, not security

- **Date:** 2026-08-17
- **Status:** Accepted
- **Context:** Req §24 requires protection design and explicitly notes it is an accidental-change deterrent.
- **Decision:** Lock formula/header/instruction cells; protect every sheet (empty password, so an informed user can unprotect); protect workbook structure; document backup/restore baseline.
- **Alternatives considered:** No protection; strong passwords (lock users out).
- **Consequences:** Accidental edits blocked; intentional edits require unprotecting; documented in delivery docs.
- **Files/components affected:** All sheets; protection contract.
- **Supersedes:** None
