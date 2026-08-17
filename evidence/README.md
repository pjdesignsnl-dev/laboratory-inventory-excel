# Evidence — Laboratory Inventory v0.1 (macro-free)

All evidence below was produced by repeatable scripts in `scripts/` against
`workbook/LabInventory_v0.1.xlsx` on this build machine
(Windows 10 IoT Enterprise LTSC 2021, build 19044; no Microsoft Office
installed). **No test on this machine is a desktop-Excel test** — see
`docs/requirements-analysis.md` §1 and decision D-004 for the verification
boundary. Desktop Excel open/recalc/visual verification is a defined
owner-side step before the contract freeze.

## Revision note (architecture review corrections, 2026-08-17)

This revision applies the four pre-freeze corrections:
1. Contract/workbook zero-drift (helper columns documented as calculated/non-authoritative; inspector compares workbook columns to contract YAML exactly).
2. D-018 usable-available stock: `Status="Available" AND (ExpiryDate blank OR ExpiryDate >= TODAY())`; expired-by-date containers excluded from stock without mutating Status; Scan blocks TakeOpen.
3. D-016 status set reduced to 6 (`Available`, `InUse`, `Expired`, `Damaged`, `Disposed`, `Missing`); `Reserved` removed.
4. D-019 dashboard views completed: most frequently used products and inventory by storage location.
5. D-013 deployment reconciled to Proposed (owner decision pending).

## Inventory exports (from `scripts/inspect_workbook.py`)

- `workbook-inventory.txt` — sheets, Tables + columns, named ranges, sampled formulas, validation rules, conditional formatting, protection.
- `formula-inventory.csv` — every formula cell (sheet, cell, formula).
- `validation-inventory.csv` — every data-validation rule.
- `table-*.csv` — each Excel Table exported as CSV (synthetic fixtures).

## Test results

- `non-vba-structural-report.txt` — structural inspection incl. workbook-vs-contract-YAML parity (49 checks, all PASS).
- `non-vba-formula-results.txt` — independent Excel-formula evaluation via the `formulas` library (55 checks, all PASS), incl. D-018 boundary tests and new dashboard views.

## Screenshots (from `scripts/render_screenshots.py`)

Representative layout previews rendered with Pillow from the workbook's actual
cells (values + fills). These are **structural previews, not desktop-Excel
screenshots**; they prove layout/cell content, not Excel rendering.

- `screenshots/01-dashboard.png`
- `screenshots/02-scan.png`
- `screenshots/03-receiving.png`
- `screenshots/04-products.png`
- `screenshots/05-containers.png`
- `screenshots/06-transactions.png`

## Checksums

```
workbook/LabInventory_v0.1.xlsx
SHA-256: 6DE5B84AE81A04F0886BD96FC35EB316AEA4218C4AA386C3065CF3E322D82470
Size:    278887 bytes
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
```

Requires `openpyxl`, `formulas`, and `pillow` (installed workspace-locally
under `.tools/pylib`, gitignored).

## Scope boundary

No VBA, Office Scripts, or other automation exists anywhere in this phase.
Desktop Excel and physical-scanner verification are deferred and are owner-side
required steps before the contract may be frozen for VBA.
