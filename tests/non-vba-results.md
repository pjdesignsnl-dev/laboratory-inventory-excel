# Non-VBA test results — Laboratory Inventory v0.1

**Workbook:** `workbook/LabInventory_v0.1.xlsx` (macro-free)
**Test date:** 2026-08-17
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

## Summary

| Phase | Checks | Result |
|---|---|---|
| B — Structural inspection (openpyxl) | 44 | **PASS** |
| C — Formula & business rules (formulas library) | 37 | **PASS** |
| D — VBA module tests | — | DEFERRED (prohibited this phase) |
| E — Integration / physical scanner | — | DEFERRED (needs desktop Excel + scanner) |

## Phase A — Architecture tests (documentation)

| Test | Expected | Actual | Result | Evidence |
|---|---|---|---|---|
| A-001 | Requirements/assumptions analysis complete | `docs/requirements-analysis.md` | PASS | doc |
| A-002 | Architecture complete (sheets, tables, columns, keys, lists, barcode, formulas, workflows, transitions, deployment) | `docs/architecture.md` | PASS | doc |
| A-003 | Status/transaction transition matrix complete | `docs/status-transition-matrix.md` | PASS | doc |
| A-004 | Decisions log updated with all material choices | `docs/decisions.md` (D-001…D-015) | PASS | doc |
| A-005 | Contract YAML drafted (`status: draft`, `vba_authorized: false`) | `schema/workbook-contract.yaml` | PASS | yaml |

## Phase B — Structural inspection (openpyxl), 44 checks

Full detail: `evidence/non-vba-structural-report.txt`; run: `python scripts/inspect_workbook.py`

| Test ID | Check | Result | Evidence |
|---|---|---|---|
| B-001 | Sheet set matches contract (9 sheets) | PASS | report |
| B-002 | `tblProducts` headers match contract (21 cols) | PASS | report |
| B-003 | `tblContainers` headers match contract (15 cols) | PASS | report |
| B-004 | `tblTransactions` headers match contract (16 cols) | PASS | report |
| B-005 | `tblSuppliers` headers match contract (8 cols) | PASS | report |
| B-006 | `tblLocations` headers match contract (5 cols) | PASS | report |
| B-007 | `tblSettings` headers match contract (3 cols) | PASS | report |
| B-008 | `tblStatusList` headers match contract (2 cols) | PASS | report |
| B-009 | `tblTransactionTypeList` headers match contract (2 cols) | PASS | report |
| B-010 | `tblExpiryClassList` headers match contract (2 cols) | PASS | report |
| B-011 | `tblScanResults` headers match contract (13 cols) | PASS | report |
| B-012 | `tblReceiveStaging` headers match contract (12 cols) | PASS | report |
| B-013 | All contract named ranges present (30) | PASS | report |
| B-014 | Products PK unique + nonblank (n=6) | PASS | report |
| B-015 | Containers PK unique + nonblank (n=20) | PASS | report |
| B-016 | Transactions PK unique + nonblank (n=31) | PASS | report |
| B-017 | Suppliers PK unique + nonblank (n=3) | PASS | report |
| B-018 | Locations PK unique + nonblank (n=6) | PASS | report |
| B-019 | Containers Barcode unique (n=20) | PASS | report |
| B-020 | All barcodes 7-digit text (leading zeros preserved) | PASS | report |
| B-021 | Containers→Products FK valid | PASS | report |
| B-022 | Products→Suppliers FK valid | PASS | report |
| B-023 | Containers→Locations FK valid | PASS | report |
| B-024 | Transactions→Containers FK valid | PASS | report |
| B-025 | Transactions→Products FK valid | PASS | report |
| B-026 | Transactions→Barcode FK valid | PASS | report |
| B-027 | Container Status values in controlled list | PASS | report |
| B-028 | TransactionType values in controlled list | PASS | report |
| B-029 | ProductType values in controlled list | PASS | report |
| B-030 | Category values in controlled list | PASS | report |
| B-031 | ExpiryDate ≥ DateReceived where both set | PASS | report |
| B-032 | Products HelperAvailableStock is a formula | PASS | report |
| B-033 | Products HelperStockClass is a formula | PASS | report |
| B-034 | Dashboard has formula stats | PASS | report |
| B-035 | Scan staging row has formulas | PASS | report |
| B-036 | Transactions append-only (all rows have IDs; no blanks) | PASS | report |
| B-037 | All 9 sheets protected | PASS | report |
| B-038 | Workbook structure protected | PASS | report |
| B-039 | Validation rules present (list + custom, see inventory) | PASS | `evidence/validation-inventory.csv` |
| B-040 | Conditional formatting present (status/expiry/duplicate) | PASS | `evidence/workbook-inventory.txt` |
| B-041 | Formula inventory exported | PASS | `evidence/formula-inventory.csv` |
| B-042 | Table CSV exports produced | PASS | `evidence/table-*.csv` |
| B-043 | Screenshots rendered | PASS | `evidence/screenshots/*.png` |
| B-044 | Workbook loads cleanly with openpyxl | PASS | inspector ran |

## Phase C — Formula & business-rule tests (independent engine), 37 checks

Full detail: `evidence/non-vba-formula-results.txt`; run: `python scripts/test_formulas.py`

Method: the workbook is loaded, every structured reference (`tblX[Col]`) and
named range is mechanically rewritten to the equivalent plain A1 range (formula
logic untouched), Tables are removed, and the `formulas` library recalculates.
Expectations derive independently from the synthetic fixtures.

| Test ID | Check | Actual | Expected | Result |
|---|---|---|---|---|
| F01-P000001 | Available stock Pipette Tips | 3 | 3 | PASS |
| F01-P000002 | Available stock Tubes | 2 | 2 | PASS |
| F01-P000003 | Available stock Ethanol | 4 | 4 | PASS |
| F01-P000004 | Available stock Methanol | 1 | 1 | PASS |
| F01-P000005 | Available stock Tris reagent | 2 | 2 | PASS |
| F01-P000006 | Available stock Gloves | 2 | 2 | PASS |
| F02-P000001 | Stock class Pipette Tips | Low | Low | PASS |
| F02-P000002 | Stock class Tubes | Low | Low | PASS |
| F02-P000003 | Stock class Ethanol | OK | OK | PASS |
| F02-P000004 | Stock class Methanol | Low | Low | PASS |
| F02-P000005 | Stock class Tris reagent | Low | Low | PASS |
| F02-P000006 | Stock class Gloves | Reorder | Reorder | PASS |
| C001 | Dashboard total available containers | 14 | 14 | PASS |
| C002-dash-expired | Dashboard expired (any status) | 1 | 1 | PASS |
| C002-dash-30 | Dashboard expiring ≤30 d | 2 | 2 | PASS |
| C002-dash-60 | Dashboard expiring 31–60 d | 0 | 0 | PASS |
| C002-dash-90 | Dashboard expiring 61–90 d | 1 | 1 | PASS |
| C003-dash-oos | Dashboard out-of-stock | 0 | 0 | PASS |
| C003-dash-reorder | Dashboard reorder | 1 | 1 | PASS |
| C003-dash-low | Dashboard low | 4 | 4 | PASS |
| C004-r34…r39 | Dashboard category counts (6) | match | match | PASS |
| F05-0/1/2 | Scan lookup known / unknown / empty (ContainerID, Status, LookupState) | match | match | PASS |
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
4. Confirm Scan sheet: type/scan `0000001` → details resolve; unknown barcode → UNKNOWN.
5. Confirm Receiving sheet computes next ContainerID/Barcode.
6. Confirm Dashboard numbers reconcile to Containers/Transactions.
7. Physical scanner acceptance (Enter suffix) per `tests/test-plan.md` Phase E.

## Worked-example coverage (Req §25)

| Example | Covered by |
|---|---|
| Box of pipette tips | P000001, C000001–C000003, C000020 (Missing) |
| Box of laboratory tubes | P000002, C000004–C000006, C000019 (Reserved) |
| Bottle of ethanol | P000003, C000007–C000009, C000018 |
| Second solvent (methanol) | P000004, C000010–C000011 |
| Reagent with batch + expiry | P000005, C000012–C000014 (expiring + expired) |
| Complete histories | 31 transactions covering Receive/TakeOpen/Return/Transfer/Dispose/MarkExpired/MarkMissing/Adjustment |

## Workbook checksum

```
SHA-256: C48F649F98F3FF1A9BF4574177705C09C0C0A75FE54D1038435E1A44C7DDE80A
Path:    workbook/LabInventory_v0.1.xlsx
Size:    278210 bytes
```

## Reproduction

```
python scripts/build_workbook.py
python scripts/inspect_workbook.py
python scripts/test_formulas.py
python scripts/render_screenshots.py
```
