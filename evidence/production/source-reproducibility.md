# Production binary ↔ source reproducibility verification

Date: 2026-08-19
Branch: `feat/non-vba-v0.1` (HEAD `1fd3e095`)
Workbook: `workbook/LabInventory_v1.0.0-production.xlsm`
SHA-256: `A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1`

## Method

Compare every module exported from the production binary
(`vba/exported-production/`) against the source-controlled modules
(`vba/modules/` and `vba/docmodules/`), normalizing for Excel's VBE
export/import behavior:

- line endings (CRLF ↔ LF),
- identifier case (VBA is case-insensitive; Excel lowercases identifiers on
  export, e.g. `MsgClass` → `msgClass`, `.Count` → `.count`),
- the VBE export header block (`VERSION 1.0 CLASS / BEGIN … END /
  Attribute VB_*`) and the `Attribute VB_Name` line that the build strips
  when injecting document-module code via `AddFromString`.

## Results

| Module | Source | Production export | Code identical |
|---|---|---|---|
| modConstants | `vba/modules/modConstants.bas` | `vba/exported-production/modConstants.bas` | **YES** |
| modWorkbookContract | … | … | **YES** |
| modUtilities | … | … | **YES** |
| modBarcodeLookup | … | … | **YES** |
| modValidation | … | … | **YES** |
| modTransactions | … | … | **YES** |
| modContainers | … | … | **YES** |
| modReceiving | … | … | **YES** |
| modScanInterface | … | … | **YES** |
| modBackup | … | … | **YES** |
| modErrorHandling | … | … | **YES** |
| modCode128 | … | … | **YES** |
| modFaultInjection | … | … | **YES** |
| ThisWorkbook (document) | `vba/docmodules/ThisWorkbook.cls` | `vba/exported-production/ThisWorkbook.cls` | **YES** |
| Scan (Sheet2, document) | `vba/docmodules/Scan.cls` | `vba/exported-production/Sheet2.cls` | **YES** |

All 13 standard modules are byte-identical after LF + case normalization
(0 mismatches). Both document modules are code-identical after stripping the
VBE export wrapper. The production binary therefore embeds **exactly** the
source-controlled code.

## Event/validation code presence (spot-check)

- `Sheet2.cls` (Scan): `Worksheet_Change`, `HandleScannedBarcode`,
  `modScanInterface` present — the scan dispatch event is wired.
- `ThisWorkbook.cls`: `Workbook_Open`, `ContractValidate`, `FailClosed`
  present — the runtime contract validation is wired.

## Reproducibility conclusion

The `scripts/build_production.ps1` pipeline is deterministic with respect to
source: identical source modules produce identical embedded code (the binary
bytes differ only by Excel's embedded metadata/timestamps, which do not
affect behavior). The committed production workbook is the pipeline's output
and is verified to match source. Combined with the immutable Git history and
recorded SHA-256, the production artifact is reproducible and versioned.
