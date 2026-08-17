# Evidence — Laboratory Inventory v0.1 (macro-free)

All evidence below was produced by repeatable scripts in `scripts/` against
`workbook/LabInventory_v0.1.xlsx` on this build machine
(Windows 10 IoT Enterprise LTSC 2021, build 19044; no Microsoft Office
installed). **No test on this machine is a desktop-Excel test** — see
`docs/requirements-analysis.md` §1 and decision D-004 for the verification
boundary. Desktop Excel open/recalc/visual verification is a defined
owner-side step before the contract freeze.

## Inventory exports (from `scripts/inspect_workbook.py`)

- `workbook-inventory.txt` — sheets, Tables + columns, named ranges, sampled formulas, validation rules, conditional formatting, protection.
- `formula-inventory.csv` — every formula cell (sheet, cell, formula).
- `validation-inventory.csv` — every data-validation rule.
- `table-*.csv` — each Excel Table exported as CSV (synthetic fixtures).

## Test results

- `non-vba-structural-report.txt` — structural inspection (44 checks, all PASS).
- `non-vba-formula-results.txt` — independent Excel-formula evaluation via the `formulas` library (37 checks, all PASS).

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
SHA-256: C48F649F98F3FF1A9BF4574177705C09C0C0A75FE54D1038435E1A44C7DDE80A
Size:    278210 bytes
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
