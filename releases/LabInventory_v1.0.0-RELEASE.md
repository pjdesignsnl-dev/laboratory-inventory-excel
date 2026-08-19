# Release Manifest — Laboratory Inventory Excel v1.0.0

**Release status:** Candidate release for review. All defined requirements and
automated/runtime acceptance tests pass with **no known defects**.

> Per project policy, this does NOT claim "100% correct". Physical-scanner
> acceptance and the external deployment choice remain **unresolved** (see
> Known limitations).

---

## 1. Artifacts

| Artifact | Path | SHA-256 |
|---|---|---|
| Operational workbook (macro-enabled) | `workbook/LabInventory_v1.0.0.xlsm` | `71D8739F0991BC074EB4489CE4A97D4891C7034640D03B2AACEB8875FC221F4E` |
| Read-only report workbook (macro-free) | `workbook/LabInventory_v1.0.0-readonly-report.xlsx` | derived from frozen base (below) |
| Macro-free base (frozen, D-022 revised) | `workbook/LabInventory_v0.1.xlsx` | `2CD1B7E6E2CB44B418134AF16F1686ADC04FECC06C9ADFCE55442BFF2BADC9E5` |
| VBA source (modules + document modules + release export) | `vba/modules/*.bas`, `vba/docmodules/*.cls`, `vba/exported-release/` | source-controlled |

## 2. Version identity

| Item | Value |
|---|---|
| Contract version | 1.0.0 (FROZEN FOR VBA) |
| Contract schema | `schema/workbook-contract.yaml` — `status: frozen`, `vba_authorized: true` |
| Application version | 1.0.0 (candidate release) |
| Excel version tested | Microsoft Excel 16.0.20228.20190 (x64), Windows |
| Excel runtime acceptance (base) | 29/29 PASS (post D-022) |
| Structural (non-VBA) | 51/51 PASS |
| Formula / business-rule (non-VBA) | 55/55 PASS |
| VBA test sweep (`Test_RunAll`) | all PASS in real Excel |
| VBA Phase F matrix + reconciliation | all PASS in real Excel |
| Fault-injection atomicity (`Test_FaultInjection`) | all PASS (5 scan-commit + 5 receive boundaries + ReceiveN batch rollback) |
| Scanner simulation (keyboard-wedge events) | PASS |
| Save/close/reopen persistence | PASS |
| Performance smoke (`Test_PerfSmoke`) | PASS (receive 20 in 1.18 s; deep lookup <1 ms; 20 lookups 8 ms) |

## 3. Git identity

| Commit | SHA (abbrev) | Purpose |
|---|---|---|
| Architecture source (pre-freeze HEAD) | `731b7a13f6` | accepted runtime 29/29 |
| **Freeze commit** | `bd65de1` | contract frozen; deployment stays proposed/unresolved |
| VBA design | `b49aeec` | `docs/vba-design.md` |
| VBA Stage 1 | `d23f2b1` | constants + runtime contract validator |
| VBA stages 2–13 + D-022 | `49ba2f7` | full module set; protection-hash fix |
| Phase F matrix | `d97c0ba` | all 9 transitions + D-018 + reconciliation |
| Phase F/G events + scanner sim | `e7c9a10` | event dispatch validated |
| Performance + regression | `8ffc1b5` | 500-container scale test |
| Release v1.0.0 (first) | `53aa89a`, `c59d6ea` | release build + final acceptance |
| **Fault-injection acceptance** | `3b9c4d2` | atomicity under injected failures (modFaultInjection) |
| **Release HEAD (candidate)** | `3b9c4d25fd` | this manifest's source tree |

## 4. Workbook contract (frozen members)

- 9 worksheets: Dashboard, Scan, Receiving, Products, Containers, Transactions,
  Suppliers, Locations, Settings.
- 11 ListObjects with exact columns and header text == column name
  (tblProducts, tblContainers, tblTransactions, tblSuppliers, tblLocations,
  tblSettings, tblStatusList, tblTransactionTypeList, tblExpiryClassList,
  tblScanResults, tblReceiveStaging).
- Named ranges: rngScanInput (Scan!D7), rngScanStatusMessage (Scan!D9),
  rngScanResultCard (Scan!D4:K34), rngReceive* (Receiving), settings names
  ExpiryWarningDays30/60/90, DefaultLocationID, DefaultStatusNewContainers.
- 6 statuses: Available, InUse, Expired, Damaged, Disposed, Missing.
- 9 transaction types: Receive, TakeOpen, Return, Transfer, Dispose,
  MarkExpired, MarkDamaged, MarkMissing, Adjustment.
- IDs: C######, P######, T########, S######, LOC####; barcode `\d{7}` text.
- D-018: usable available = Status="Available" AND (no expiry OR expiry >= today);
  expired-by-date Available never silently mutated; TakeOpen blocked.

## 5. VBA module inventory (source-controlled)

| Module | Responsibility |
|---|---|
| `modConstants` | All frozen constants (sheets, tables, columns, statuses, txns, formats) |
| `modWorkbookContract` | Runtime contract validator (fail-closed), diagnostics |
| `modUtilities` | Clock (`GetNow`), operator capture, state save/restore, sheet protection helpers, table map, barcode normalize |
| `modBarcodeLookup` | Stateless text-barcode scan → row; duplicate count; staging population |
| `modValidation` | Transition matrix + D-018 expiry semantics; required-field/FK checks |
| `modTransactions` | T######## ID generation; snapshot builder; atomic append; rollback-removal |
| `modContainers` | C######/barcode ID generation (PadID); state capture/apply/rollback; AddContainer |
| `modReceiving` | ReceiveOne / ReceiveN (all-or-nothing batch) |
| `modScanInterface` | Scan dispatch, commit pipeline, m_busy re-entrancy guard, reset/focus |
| `modBackup` | Timestamped SaveCopyAs backup (backup-required-fails-stops) |
| `modErrorHandling` | Error classification, operator messages, diagnostics log |
| `modCode128` | Code 128B pattern generation (printer-independent labels) |
| `modFaultInjection` | Deterministic one-shot fault injection at mutation boundaries (test-only) |
| `modTestHooks` | Test hooks (all suites) |
| `ThisWorkbook` | Workbook_Open contract validation (fail-closed) |
| `Scan` (Sheet2) | Worksheet_Change → dispatch only (no business rules) |

## 6. Test totals (real desktop Excel)

| Suite | Totals |
|---|---|
| Excel runtime acceptance (macro-free base) | 29/29 PASS |
| Structural inspection | 51/51 PASS |
| Formula engine / business rules | 55/55 PASS |
| VBA `Test_RunAll` (contract/drift/lookup/transitions/IDs/receive/Code128/operator/backup) | all PASS (~480 ms) |
| VBA `Test_PhaseF` (9-transition matrix, D-018 commit block, atomicity, dashboard reconciliation) | all PASS (~2 s) |
| VBA `Test_FaultInjection` (5 scan-commit boundaries + 5 receive boundaries + ReceiveN batch rollback, full invariant checks) | all PASS (~4.5 s) — see `evidence/vba/atomicity-fault-injection.md` |
| Scanner simulation (keyboard-wedge event → staging) | PASS |
| Save/close/reopen persistence | PASS |
| Performance (500 containers) | receive 500 OK (58.6 s); deep lookup 12 ms; 500 lookups 3039 ms |
| Performance smoke (`Test_PerfSmoke`) | receive 20 OK (1.18 s); deep lookup <1 ms; 20 lookups 8 ms; cleanup restores table |

## 7. Security / macro settings (AccessVBOM)

- **Before development:** `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM` was **unset** (0 = default: programmatic VBA project access disabled).
- **During development (module import/export automation):** set to `1`.
- **Restore status (2026-08-19):** after all VBA import/export work finished, restored to **0/unset** and the final `.xlsm` was re-verified to open and operate normally with AccessVBOM disabled (see `evidence/vba/release-acceptance-log.txt`).
- **Macro security:** do NOT instruct users to "enable all macros". Recommended: Trusted Location for the operational file, or digitally sign the VBA project. Physical-scanner and deployment decisions remain with the owner.

## 8. Known limitations

1. **Physical barcode scanner not available** in the test environment. Scanner behavior was validated by keyboard-wedge simulation (typing barcode + Enter into Scan!D7 with events enabled). A real Code 128 scanner must be accepted on-site.
2. **Deployment option unresolved** (D-013 Proposed): `writer_model: single_writer`, `reader_model: read_only_elsewhere`, `deployment_option: unresolved`. The read-only report workbook covers the read-only viewer case.
3. **Remaining-volume / piece tracking** intentionally out of scope (v1).
4. **Batch receive throughput** ~117 ms/row in real Excel (dominated by Excel table formula auto-fill); acceptable for lab receiving. Single scans are instant (~6–12 ms).
5. **Code 128 labels**: the module generates the Code 128B pattern string; a matching TrueType Code 128 font is required on the printer machine (printer-independent by design).
6. **Undo**: transactions are append-only and container state is mutated by VBA; Excel Undo does not reverse a committed transaction (by design; corrections use Adjustment).

## 9. Unresolved / next steps

- On-site physical scanner acceptance (model/config to record in the manifest when available).
- Owner decision on deployment option (single-file vs. multi-user), then finalize the deployment recommendation.
- Optional: digital signature for the VBA project before production distribution.
