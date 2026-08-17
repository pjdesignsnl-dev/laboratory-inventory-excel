# Evidence — Laboratory Inventory v0.1 (macro-free)

All evidence below was produced by repeatable scripts in `scripts/` against
`workbook/LabInventory_v0.1.xlsx` on this build machine.

**Environment (as of 2026-08-17):** Windows 10 IoT Enterprise LTSC 2021
(build 19044), **Microsoft Excel 16.0.20228.20190 (x64) is installed** and was
used for the authoritative runtime acceptance in `evidence/excel-runtime/`.
Static openpyxl/formulas tests remain supplementary.

## Excel runtime acceptance (authoritative, `evidence/excel-runtime/`)

- `excel-runtime-results.txt` — **29/29 PASS** using the actual installed
  Microsoft Excel via COM: opens without repair, no VBA project, 9 sheets,
  all Tables, named ranges, no formula errors after full rebuild, Dashboard
  reconciliation, Scan lookups (0000001 FOUND / 9999999 UNKNOWN / 0000021
  TakeOpen blocked D-018), Receiving next-ID, protection, validation.
- `recalculation-results.txt` — full-rebuild + validation-copy save/reopen/
  recalc: no corruption, no formula drift.
- `workbook-object-inventory.txt` — worksheets, Tables, named ranges as seen
  by Excel COM.
- `screenshots/*.pdf` — PDF exports rendered by Microsoft Excel itself.
- `environment.txt` — OS/Excel/COM details.
- `LabInventory_v0.1-validation-copy.xlsx` — the reopened/recalculated copy.

## Revision note (architecture review corrections, 2026-08-17)

This revision applies the four pre-freeze corrections:
1. Contract/workbook zero-drift (helper columns documented as calculated/non-authoritative; inspector compares workbook columns to contract YAML exactly).
2. D-018 usable-available stock: `Status="Available" AND (ExpiryDate blank OR ExpiryDate >= TODAY())`; expired-by-date containers excluded from stock without mutating Status; Scan blocks TakeOpen.
3. D-016 status set reduced to 6 (`Available`, `InUse`, `Expired`, `Damaged`, `Disposed`, `Missing`); `Reserved` removed.
4. D-019 dashboard views completed: most frequently used products and inventory by storage location.
5. D-013 deployment reconciled to Proposed (owner decision pending).

**Excel-runtime defects found and fixed (2026-08-17):**
- Table header text now equals the column name (e.g., `ExpiryDate`, not
  `Expiry Date`) so Excel structured references resolve — previously they
  evaluated to `#REF!` in real Excel (static/formulas tests had masked this).
- Receiving interface formulas now use direct `$D$7` references instead of the
  `rngReceiveProductID` named range, so Excel's dynamic-array engine does not
  inject implicit-intersection and break the next-ID display.
- Receiving/Scan text entry forced via `@` number format in the runtime test.

## Inventory exports (from `scripts/inspect_workbook.py`)

- `workbook-inventory.txt` — sheets, Tables + columns, named ranges, sampled formulas, validation rules, conditional formatting, protection.
- `formula-inventory.csv` — every formula cell (sheet, cell, formula).
- `validation-inventory.csv` — every data-validation rule.
- `table-*.csv` — each Excel Table exported as CSV (synthetic fixtures).

## Test results

- `non-vba-structural-report.txt` — structural inspection incl. workbook-vs-contract-YAML parity (51 checks, all PASS).
- `non-vba-formula-results.txt` — independent Excel-formula evaluation via the `formulas` library (55 checks, all PASS), incl. D-018 boundary tests and new dashboard views.

## Screenshots (from `scripts/render_screenshots.py`)

Representative layout previews rendered with Pillow from the workbook's actual
cells (values + fills). These are **structural previews, not desktop-Excel
screenshots**; they prove layout/cell content, not Excel rendering. The
authoritative Excel-rendered screenshots are the PDFs in
`evidence/excel-runtime/screenshots/`.

- `screenshots/01-dashboard.png`
- `screenshots/02-scan.png`
- `screenshots/03-receiving.png`
- `screenshots/04-products.png`
- `screenshots/05-containers.png`
- `screenshots/06-transactions.png`

## Checksums

```
workbook/LabInventory_v0.1.xlsx
SHA-256: C3D27FE82840833459674690DA250EC1500E99CC9337E52F4A7C4FAF81ED787A
```

Regenerate with:

```
Get-FileHash workbook\LabInventory_v0.1.xlsx -Algorithm SHA256
```

## Reproduction

```
python scripts/build_workbook.py
python scripts/inspect_workbook.py
python scripts/test_formulas.py
python scripts/render_screenshots.py
& scripts/excel_runtime_test.ps1   # requires installed Microsoft Excel
```

Requires `openpyxl`, `formulas`, and `pillow` (workspace-local `.tools/pylib`).

## Scope boundary

No VBA, Office Scripts, or other automation exists anywhere in this phase.
The Excel runtime acceptance passed (29/29); the contract is still **NOT
frozen** and waits for owner architecture-freeze authorization before any VBA.
