# Non-VBA test results — Laboratory Inventory v0.1

**Workbook:** `workbook/LabInventory_v0.1.xlsx` (macro-free)
**Test date:** 2026-08-17 (rev. 2026-08-17 — architecture-review corrections + Excel runtime acceptance)
**Build machine:** Windows 10 IoT Enterprise LTSC 2021 (21H2), build 19044, AMD64
**Microsoft Excel installed:** **Yes** — 16.0.20228.20190 (x64), used for the authoritative runtime acceptance in `evidence/excel-runtime/`.

> **Verification model (decision D-004, superseded by runtime).** The static
> openpyxl/formulas tests were the only option before Excel was installed. With
> Microsoft Excel now installed, the **Excel runtime acceptance
> (evidence/excel-runtime/excel-runtime-results.txt, 29/29 PASS) is the
> authoritative check**; the openpyxl/formulas suites remain as supplementary
> regression evidence.

## Runtime acceptance summary (Microsoft Excel COM, 29/29 PASS)

Full detail: `evidence/excel-runtime/excel-runtime-results.txt`; run: `& scripts/excel_runtime_test.ps1`

| Check | Result |
|---|---|
| Workbook opens without repair (xlsx, format 51) | PASS |
| No VBA project (HasVBProject=False) | PASS |
| All 9 expected worksheets exist | PASS |
| All ListObjects/Tables present per contract | PASS |
| Named ranges resolve (bounded + column-lists by name) | PASS |
| No formula errors after CalculateFullRebuild | PASS |
| Dashboard total available = 15 | PASS |
| Dashboard expired = 2 / low = 4 / reorder = 1 | PASS |
| Dashboard frequently-used (rows 43–48) = 1,1,1,1,0,0 | PASS |
| Dashboard inventory-by-location (rows 52–57) = 5,0,2,8,0,0 | PASS |
| Scan 0000001 → FOUND / C000001 / Available | PASS |
| Scan 9999999 → UNKNOWN | PASS |
| Scan 0000021 → TakeOpen BLOCKED (expired by date, D-018) | PASS |
| P000005 usable stock = 2 (expired-by-date excluded) | PASS |
| Receiving next ContainerID = C000022 / Barcode = 0000022 | PASS |
| All 9 sheets protected; workbook structure protected | PASS |
| Data validation present (7 probed cells) | PASS |
| Validation copy saved, reopened, recalculated: no errors, no drift | PASS |

## Defects found by the Excel runtime and fixed (macro-free only)

1. **Structured-reference `#REF!`:** Excel Table header text was the display
   text ("Expiry Date") while formulas used the column name (`ExpiryDate`), so
   `tblContainers[ExpiryDate]` evaluated to `#REF!` in real Excel. Fixed: header
   text now equals the column name exactly. (Static/formulas tests had masked
   this because they rewrote references by schema name.)
2. **Receiving next-ID broken by implicit intersection:** formulas using the
   `rngReceiveProductID` named range were mangled by Excel's dynamic-array
   `@` operator. Fixed: Receiving formulas now use direct `$D$7` references
   (same pattern as the verified Scan sheet).
3. **Text entry coercion:** COM wrote `0000001` as a number unless the cell
   format was forced to text. The runtime test now forces `@` before writing
   barcodes/IDs.

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
| R — Microsoft Excel runtime acceptance (authoritative) | 29 | **PASS** |
| B — Structural inspection (openpyxl) incl. contract-YAML parity | 51 | **PASS** |
| C — Formula & business rules (formulas library) incl. D-018 + dashboard views | 55 | **PASS** |
| D — VBA module tests | — | DEFERRED (prohibited this phase) |
| E — Integration / physical scanner | — | DEFERRED (needs physical scanner) |

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
SHA-256: C3D27FE82840833459674690DA250EC1500E99CC9337E52F4A7C4FAF81ED787A
Path:    workbook/LabInventory_v0.1.xlsx
```

## Reproduction

```
python scripts/build_workbook.py
python scripts/inspect_workbook.py
python scripts/test_formulas.py
python scripts/render_screenshots.py
& scripts/excel_runtime_test.ps1   # requires installed Microsoft Excel
```
