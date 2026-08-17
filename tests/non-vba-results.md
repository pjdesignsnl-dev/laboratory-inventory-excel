# Non-VBA test results — Laboratory Inventory v0.1

**Workbook:** `workbook/LabInventory_v0.1.xlsx` (macro-free)
**Test date:** 2026-08-17 (rev. 2026-08-17 — architecture-review corrections)
**Build machine:** Windows 10 IoT Enterprise LTSC 2021 (21H2), build 19044, AMD64
**Excel installed on test machine:** **No** (no Microsoft Office/LibreOffice/WPS found)

> **Verification boundary (decision D-004).** Because no desktop Excel exists on
> the build machine, tests here are (1) openpyxl structural inspection and
> (2) independent Excel-formula evaluation with the `formulas` library on a
> mechanically flattened copy of the workbook (structured references rewritten
> to plain ranges; formula logic unchanged). These are genuine independent
> checks but are **not** desktop Excel tests. Desktop Excel open/recalc/visual
> verification and physical-scanner acceptance remain owner-side steps before
> the contract freeze — they are listed as `DEFERRED` below.

## Revision summary (2026-08-17)

Four pre-freeze corrections applied (see `docs/decisions.md` D-013, D-016–D-019):

1. **Contract/workbook drift eliminated** — helper columns documented in contract as calculated/non-authoritative; inspector now verifies exact workbook-vs-contract-YAML column parity (zero drift).
2. **D-018 usable available stock** — `Status="Available" AND (ExpiryDate blank OR ExpiryDate >= TODAY())`; expired-by-date containers excluded from stock/reorder without mutating Status; Scan blocks TakeOpen for them.
3. **D-016 status set reduced to 6** — `Reserved` removed everywhere.
4. **D-019 dashboard views added** — most frequently used products and inventory by storage location.
5. **D-013 reconciled to Proposed** — deployment is an unresolved owner decision.

## Summary

| Phase | Checks | Result |
|---|---|---|
| B — Structural inspection (openpyxl) incl. contract-YAML parity | 51 | **PASS** |
| C — Formula & business rules (formulas library) incl. D-018 + dashboard views | 55 | **PASS** |
| D — VBA module tests | — | DEFERRED (prohibited this phase) |
| E — Integration / physical scanner | — | DEFERRED (needs desktop Excel + scanner) |

## Phase A — Architecture tests (documentation)

| Test | Expected | Actual | Result | Evidence |
|---|---|---|---|---|
| A-001 | Requirements/assumptions analysis complete (rev. incl. D-016/D-018) | `docs/requirements-analysis.md` | PASS | doc |
| A-002 | Architecture complete (sheets, tables, columns, keys, lists, barcode, formulas incl. D-018, workflows, transitions, deployment) | `docs/architecture.md` | PASS | doc |
| A-003 | Status/transaction transition matrix complete (6-status model) | `docs/status-transition-matrix.md` | PASS | doc |
| A-004 | Decisions log updated (D-001…D-019; D-013 → Proposed) | `docs/decisions.md` | PASS | doc |
| A-005 | Contract YAML drafted (`status: draft`, `vba_authorized: false`, contract 0.2.0) | `schema/workbook-contract.yaml` | PASS | yaml |

## Phase B — Structural inspection (openpyxl), 51 checks

Full detail: `evidence/non-vba-structural-report.txt`; run: `python scripts/inspect_workbook.py`

| Test ID | Check | Result | Evidence |
|---|---|---|---|
| B-001 | Sheet set matches contract (9 sheets) | PASS | report |
| B-002 | `tblProducts` headers match contract (21 cols incl. helpers) | PASS | report |
| B-003 | `tblContainers` headers match contract (15 cols incl. helpers) | PASS | report |
| B-004 | `tblTransactions` headers match contract (16 cols) | PASS | report |
| B-005 | `tblSuppliers` headers match contract (8 cols) | PASS | report |
| B-006 | `tblLocations` headers match contract (5 cols) | PASS | report |
| B-007 | `tblSettings` headers match contract (3 cols) | PASS | report |
| B-008 | `tblStatusList` headers match contract (2 cols) — 6 statuses | PASS | report |
| B-009 | `tblTransactionTypeList` headers match contract (2 cols) | PASS | report |
| B-010 | `tblExpiryClassList` headers match contract (2 cols) | PASS | report |
| B-011 | `tblScanResults` headers match contract (13 cols) | PASS | report |
| B-012 | `tblReceiveStaging` headers match contract (12 cols) | PASS | report |
| B-013 | **Workbook vs contract YAML zero drift** | PASS | report |
| B-014 | **Helper/calculated columns documented as non-authoritative in contract** | PASS | report |
| B-015 | All contract named ranges present (30) | PASS | report |
| B-016 | Products PK unique + nonblank (n=6) | PASS | report |
| B-017 | Containers PK unique + nonblank (n=21) | PASS | report |
| B-018 | Transactions PK unique + nonblank (n=31) | PASS | report |
| B-019 | Suppliers PK unique + nonblank (n=3) | PASS | report |
| B-020 | Locations PK unique + nonblank (n=6) | PASS | report |
| B-021 | Containers Barcode unique (n=21) | PASS | report |
| B-022 | All barcodes 7-digit text (leading zeros preserved) | PASS | report |
| B-023 | Containers→Products FK valid | PASS | report |
| B-024 | Products→Suppliers FK valid | PASS | report |
| B-025 | Containers→Locations FK valid | PASS | report |
| B-026 | Transactions→Containers FK valid | PASS | report |
| B-027 | Transactions→Products FK valid | PASS | report |
| B-028 | Transactions→Barcode FK valid | PASS | report |
| B-029 | Container Status values in controlled list (6 values) | PASS | report |
| B-030 | TransactionType values in controlled list | PASS | report |
| B-031 | ProductType values in controlled list | PASS | report |
| B-032 | Category values in controlled list | PASS | report |
| B-033 | ExpiryDate ≥ DateReceived where both set | PASS | report |
| B-034 | Products HelperAvailableStock is a formula | PASS | report |
| B-035 | Products HelperStockClass is a formula | PASS | report |
| B-036 | Dashboard has formula stats | PASS | report |
| B-037 | Scan staging row has formulas | PASS | report |
| B-038 | Transactions append-only (all rows have IDs; no blanks) | PASS | report |
| B-039 | All 9 sheets protected | PASS | report |
| B-040 | Workbook structure protected | PASS | report |
| B-041 | Validation rules present (list + custom, see inventory) | PASS | `evidence/validation-inventory.csv` |
| B-042 | Conditional formatting present (status/expiry/duplicate) | PASS | `evidence/workbook-inventory.txt` |
| B-043 | Formula inventory exported | PASS | `evidence/formula-inventory.csv` |
| B-044 | Table CSV exports produced | PASS | `evidence/table-*.csv` |
| B-045 | Screenshots rendered | PASS | `evidence/screenshots/*.png` |
| B-046 | Workbook loads cleanly with openpyxl | PASS | inspector ran |
| B-047 | Date cells are real dates (not text) | PASS | report |
| B-048 | No `Reserved` anywhere in workbook lists/fixtures (D-016) | PASS | report + CSVs |
| B-049 | Dashboard has 'most frequently used products' section (D-019) | PASS | report |
| B-050 | Dashboard has 'inventory by storage location' section (D-019) | PASS | report |
| B-051 | Dashboard has formula stats (source reconciliation) | PASS | report |

## Phase C — Formula & business-rule tests (independent engine), 55 checks

Full detail: `evidence/non-vba-formula-results.txt`; run: `python scripts/test_formulas.py`

Method: the workbook is loaded, every structured reference (`tblX[Col]`) and
named range is mechanically rewritten to the equivalent plain A1 range (formula
logic untouched), Tables are removed, and the `formulas` library recalculates.
Expectations derive independently from the synthetic fixtures.

| Test ID | Check | Actual | Expected | Result |
|---|---|---|---|---|
| F01-P000001…006 | Available stock per product (D-018) | 3/3/4/1/2/2 | same | PASS |
| F02-P000001…006 | Stock class per product | Low/Low/OK/Low/Low/Reorder | same | PASS |
| C001 | Dashboard total available containers (usable) | 15 | 15 | PASS |
| C002-dash-expired | Dashboard expired (any status, by date) | 2 | 2 | PASS |
| C002-dash-30/60/90 | Dashboard expiring bands | 2/0/1 | 2/0/1 | PASS |
| C003-dash-oos/reorder/low | Dashboard stock-class counts | 0/1/4 | 0/1/4 | PASS |
| C004-r34…r39 | Dashboard category counts (6) | match | match | PASS |
| C005-r43…r48 | **Dashboard frequently-used products (TakeOpen counts)** | match | match | PASS |
| C006-r52…r57 | **Dashboard inventory by storage location (usable)** | match | match | PASS |
| D018-a | Available + expiring tomorrow counts today | 1 | 1 | PASS |
| D018-b | Available + expired-by-date excluded | 0 | 0 | PASS |
| D018-c | Status not silently changed (C000021 stays Available) | Available | Available | PASS |
| D018-scan-actions/block/ok | Scan TakeOpen blocking for expired-by-date | match | match | PASS |
| F05-0/1/2 | Scan lookup known / unknown / empty | match | match | PASS |
| F06-dup | Duplicate barcode flag/state (mirror) | DUPLICATE | DUPLICATE | PASS |

**Engine limitation note:** the `formulas` library returns `#VALUE!` for
`COUNTIF(range, text_criteria)` — a library quirk, not a workbook defect
(COUNTIF with text is standard Excel). The lookup path therefore uses
`MATCH`/`INDEX` (which the engine evaluates correctly), and duplicate
detection is validated via an exact mirror of the workbook's COUNTIF logic.

## Phase D — VBA module tests

**NOT RUN — VBA is explicitly prohibited in this phase** (`docs/INITIAL_TASK.md`;
`AGENTS.md` non-negotiable rule). No `.bas`, `.cls`, `.frm`, embedded macros,
Office Scripts, or other automation exist anywhere in the repository.

## Phase E — Integration and physical acceptance

**DEFERRED.** Requires Windows desktop Excel and a physical USB keyboard-wedge
scanner. Owner-side steps before contract freeze:

1. Open `workbook/LabInventory_v0.1.xlsx` in Excel 2021/2024 or Microsoft 365.
2. Confirm all formulas recalculate with no `#VALUE!`/`#REF!` errors.
3. Confirm validation dropdowns, conditional formatting, and sheet protection behave.
4. Confirm Scan sheet: type/scan `0000001` → details resolve; scan `0000021` (expired-by-date Available) → TakeOpen blocked, MarkExpired offered; unknown barcode → UNKNOWN.
5. Confirm Receiving sheet computes next ContainerID/Barcode.
6. Confirm Dashboard numbers reconcile to Containers/Transactions, including frequently-used and by-location views.
7. Physical scanner acceptance (Enter suffix) per `tests/test-plan.md` Phase E.

## Worked-example coverage (Req §25)

| Example | Covered by |
|---|---|
| Box of pipette tips | P000001, C000001–C000003, C000020 (Missing) |
| Box of laboratory tubes | P000002, C000004–C000006, C000019 (Available, no expiry) |
| Bottle of ethanol | P000003, C000007–C000009, C000018 |
| Second solvent (methanol) | P000004, C000010–C000011 |
| Reagent with batch + expiry | P000005, C000012–C000014 (expiring + expired), C000021 (expired-by-date, D-018) |
| Complete histories | 31 transactions covering Receive/TakeOpen/Return/Transfer/Dispose/MarkExpired/MarkMissing |

## Workbook checksum

```
SHA-256: 6DE5B84AE81A04F0886BD96FC35EB316AEA4218C4AA386C3065CF3E322D82470
Path:    workbook/LabInventory_v0.1.xlsx
Size:    278887 bytes
```

## Reproduction

```
python scripts/build_workbook.py
python scripts/inspect_workbook.py
python scripts/test_formulas.py
python scripts/render_screenshots.py
```
