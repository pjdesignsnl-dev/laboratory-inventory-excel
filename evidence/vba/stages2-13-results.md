# VBA Stage 2–13 Integration Results (real desktop Excel)

Date: 2026-08-18 (Excel 16.0.20228.20190 x64, Windows)
Candidate: `workbook/LabInventory_v1.0-candidate.xlsm` (built from hash-free-protection `LabInventory_v0.1.xlsx`, SHA-256 `2CD1B7E6…`)

## Environment / harness notes

- COM automation drives `Application.Run` against pre-embedded test hooks; events disabled for the test session.
- The VBE window is hidden at build-save time; an open VBE blocks `Application.Run` under COM (documented in this file's history).
- Sheet protection uses no password hash (D-022) so VBA `Unprotect` works.

## Test hook sweep (`modTestHooks.Test_RunAll`) — 378 ms, all PASS

| Test | Result |
|---|---|
| Contract check (`ContractValidate`) | `contract:OK diag=[]` |
| Contract drift (named-range delete/re-add) | `drift:before=FAIL after=OK` |
| Barcode lookup FOUND row 1 | `lookupbybarcode-end:OK row=1` |
| Barcode lookup UNKNOWN (9999999) | `unknown-end:OK` |
| Barcode lookup EMPTY | `lookup-empty:OK` |
| Barcode lookup INVALID (abc1234) | `lookup-invalid:OK` |
| Transition Available→TakeOpen | `t-avail-takeopen:OK` |
| Transition Available expired-by-date→TakeOpen BLOCKED (D-018) | `t-expiredby-date-takeopen:OK` |
| Transition InUse→Return | `t-inuse-return:OK` |
| Transition Disposed→TakeOpen BLOCKED | `t-disposed-takeopen:OK` |
| Transition Expired→Dispose | `t-expired-dispose:OK` |
| TransactionID format/unique | `txnid-format:OK tid=T00000032` `txnid-unique:OK` |
| ReceiveOne | `receive-one:OK cid=C000022` `receive-one-cid:OK` |
| ReceiveN (3) unique IDs | `receive-n:OK count=3` `receive-n-ids:C000023,C000024,C000025` `receive-n-unique:OK` |
| Code128 pattern | `code128:OK len=10` |
| Operator capture | `operator:OK op=Q` |
| Backup timestamped copy | `backup:OK path=...LabInventory_backup_20260818_165901.xlsm` |

## Scan-commit pipeline (`modTestHooks.Test_CommitTakeOpen`) — 554 ms, PASS

- `commit-diag:staging-barcode=[0000001] lookupstate=0` (staging stores barcode as text)
- `commit-takeopen:OK msg=OK: TakeOpen committed (T00000032).`

## Key fixes found during this phase

1. **Module-qualified `Enum` types in `Dim` declarations deadlock VBA/COM** when the procedure is invoked via `Application.Run`. Replaced with `Long` + named constants.
2. **`CreateObject("Scripting.Dictionary")` deadlocks** in this harness; replaced with a stateless per-call scan of the frozen `tblContainers[Barcode]` column.
3. **Module-level dynamic arrays (`Private m_barcodes() As String`) deadlock** on first access from a Run-invoked procedure; removed all module-level array state (stateless lookups).
4. **openpyxl `password="CE4B"` hash on sheet protection deadlocks VBA `Worksheet.Unprotect`** under COM (it prompts for a password). Removed the hash via `scripts/build_workbook.py::_strip_protection_password` (D-022); protection deterrent preserved; 29/29 runtime acceptance still passes.
5. **Barcode cells must be text**: staging writes now force `NumberFormat="@"` for Barcode/ContainerID/ProductID; `CommitAction` normalizes the read barcode.
6. **`Format$(n, "C000000")` locale date-misparse**: replaced with explicit `PadID` zero-padding (IDs now render `C000022` correctly).

## Non-VBA regression after protection change

- Structural inspection: 51/51 PASS (sheets still protected, structure protected).
- Formula/business-rule: 55/55 PASS.
- Real-Excel runtime acceptance: 29/29 PASS.

## Remaining (next phases)

- Worksheet event wiring for Scan/Receiving document modules (code injected; needs event-driven test).
- Code 128 label rendering integration and scanner acceptance (physical scanner not available — documented as pending).
- Full test matrix expansion, final candidate release build, release manifest.
