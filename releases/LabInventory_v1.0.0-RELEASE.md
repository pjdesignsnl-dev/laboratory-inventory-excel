# Release Manifest — Laboratory Inventory Excel v1.0.0

**Release status:** Production candidate — software acceptance passes with no
known defects. Deployment is finalized (D-023 Accepted); the remaining open
production gate(s): **physical barcode-scanner acceptance**, **physical label
print**, **deployment share provisioning**, and **initial inventory import**
(owner/hardware dependencies — see Known limitations).

> Per project policy, this does NOT claim "100% correct", and it is not yet
> "production ready" because physical scanner acceptance has not been
> performed (no scanner available). Once the open gates pass, update this
> status to "Production approved".

---

## 1. Artifacts

| Artifact | Path | SHA-256 |
|---|---|---|
| Operational workbook (macro-enabled, full test suite) | `workbook/LabInventory_v1.0.0.xlsm` | `DF26793793CA7B05C6217A674E4D9D26441EE07D7B0E1133B2B01CBFFB32692F` |
| **Production operational workbook** (clean, no fixtures, no test hooks) | `workbook/LabInventory_v1.0.0-production.xlsm` | `A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1` |
| **Read-only report workbook** (macro-free, clean) | `workbook/LabInventory_v1.0.0-readonly-report.xlsx` | `DF18E9D782D5B581C17A789B831C8FC26FFC423DE0A5A09308DDC81A4313B601` |
| Macro-free base (frozen, D-022 revised) | `workbook/LabInventory_v0.1.xlsx` | `2CD1B7E6E2CB44B418134AF16F1686ADC04FECC06C9ADFCE55442BFF2BADC9E5` |
| VBA source (modules + document modules + exports) | `vba/modules/*.bas`, `vba/docmodules/*.cls`, `vba/exported-release/`, `vba/exported-production/` | source-controlled |

## 2. Version identity

| Item | Value |
|---|---|
| Contract version | 1.0.0 (FROZEN FOR VBA) |
| Contract schema | `schema/workbook-contract.yaml` — `status: frozen`, `vba_authorized: true`; deployment `accepted` (D-023) |
| Application version | 1.0.0 (production candidate) |
| Excel version tested | Microsoft Excel 16.0.20228.20190 (x64), Windows |
| Excel runtime acceptance (base) | 29/29 PASS (post D-022) |
| Structural (non-VBA) | 51/51 PASS |
| Formula / business-rule (non-VBA) | 55/55 PASS |
| VBA test sweep (`Test_RunAll`) | all PASS in real Excel |
| VBA Phase F matrix + reconciliation | all PASS in real Excel |
| Fault-injection atomicity (`Test_FaultInjection`) | all PASS (5 scan-commit + 5 receive boundaries + ReceiveN batch rollback) |
| Scanner simulation (keyboard-wedge events) | PASS |
| Save/close/reopen persistence | PASS |
| Backup/restore drill | PASS (evidence/production/backup-restore-test.md) |
| Read-only report acceptance | PASS (macro-free, protected, no write path) |
| Production engine smoke (import → receive → TakeOpen → Return → persist) | PASS |
| Performance smoke (`Test_PerfSmoke`) | PASS (receive 20 in ~1.3 s; deep lookup <1 ms) |
| Production binary verify (empty tables, lists 6/9, contract OK, AccessVBOM 0) | PASS |

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
| Release v1.0.0 (rebuilt) | `c065112` | release rebuild + manifest update (new SHA `71D8739F…`) |
| AccessVBOM restore | `adabb53` | AccessVBOM=0; release verified with it disabled |
| **Production source HEAD** | `adabb53ee3a` | exact production source tree (pre-production-prep) |
| Production prep | `7adfa4e` | D-023 deployment decision, clean production workbook, text-format fix, engine smoke |
| Production finalize (docs/drill/report/templates/manifest) | `f0a0f68` | production docs, backup/restore drill, read-only report, import templates, label layout |
| AccessVBOM final + verify | `f9225e1` | AccessVBOM restored to 0; production re-verified |
| Deployment-location smoke | `b2b09c5` | Phase 11: smoke from simulated share location (second writer not permitted) |
| Transient artifacts cleanup | `e028ec2` | deployment-location smoke xlsm copies gitignored |
| Phase 12/13 review + integrity | `b3155e4` | UI review (no test hooks) + final security/integrity check PASS |
| Manifest identity consistency | `95870ca` | final production commit chain + acceptance HEAD recorded |
| **Final production commit (this manifest's identity)** | `95870ca6` | exact final production source tree (current HEAD) |

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
- **During development (module import/export automation):** set to `1` (temporarily, per build).
- **Restore status (2026-08-19):** restored to **0/unset** and the final `.xlsm` verified to open and operate normally with AccessVBOM disabled (see `evidence/vba/release-acceptance-log.txt`). **Production requires AccessVBOM = 0.**
- **Macro trust model (production):** no code-signing certificate is available in the environment (checked 2026-08-19), so a **controlled Trusted Location scoped to the production share folder** is the documented macro-trust mechanism (see `docs/scanner-configuration.md`). Users are NOT told to "enable all macros" or lower Trust Center security. No self-signed fake trust is presented as production-grade.
- **Mark-of-the-Web:** files transferred from the repo/downloads may be flagged; verify SHA-256 against this manifest and unblock (Properties → Unblock) or open from the Trusted Location. The trusted-location master on the writer PC is not subject to MOTW blocking.

## 8. Deployment (D-023, Accepted — supersedes D-013 Proposed)

- **Option A/B:** authoritative `LabInventory_v1.0.0-production.xlsm` on a private network share (`\\<LAB-SERVER>\Inventory\`); **only the dedicated writing PC** has write permission; all other PCs use the macro-free read-only report.
- Writer model: single writer; reader model: read-only elsewhere; **no multi-user editing**.
- File locking: Excel lock on the SMB master + NTFS write-only-for-writer.
- Versioning: timestamped backups (see `docs/backup-recovery.md`) + Git.
- Backup: `LABINV_BACKUP_FOLDER` outside the master share; retention daily 14 / weekly 8 / monthly 12.
- Offline/network-loss: master opens only from the share; no offline local copy (split-brain avoided); report is a snapshot.
- Read-only report refresh: regenerated by the writer PC after operations.
- Option C (OneDrive/SharePoint) rejected: no M365 business tenant/SharePoint provisioned. Option D (Google Drive) rejected for the writer path (backup/archive only).
- **Owner action required:** provision the share with the permission model, then run the go-live import.

## 9. Known limitations / open gates

1. **Physical barcode scanner acceptance — OPEN (hardware not available).** Scanner behavior validated by keyboard-wedge simulation only; a real Code 128 scanner must be accepted on-site (`evidence/production/scanner-acceptance.md`).
2. **Physical label print — OPEN.** Layout validated in PDF (`evidence/production/label-layout-sample.pdf`); a licensed Code 128 TrueType font + label printer are required on-site.
3. **Deployment share provisioning — OPEN (owner).** `\\<LAB-SERVER>\Inventory\` + permissions per D-023.
4. **Initial inventory import — OPEN (owner data).** Templates + QA report in `templates/`; go-live timestamp starts the audit trail; no audit events are manufactured.
5. Remaining-volume / piece tracking intentionally out of scope (v1).
6. Batch receive throughput ~117 ms/row (Excel table formula auto-fill); single scans instant (~6–12 ms).
7. Undo: committed transactions are not undoable by Excel Undo (by design; corrections use Adjustment).

## 10. Unresolved / next steps

- Run the on-site scanner + label-print acceptance; update `evidence/production/scanner-acceptance.md` and this manifest.
- Provision the deployment share (D-023) and run the initial import (`docs/initial-data-import.md`).
- Optional later: digital signature for the VBA project if a suitable certificate becomes available.
- When all open gates pass, set this manifest's status to **Production approved** and report the production-ready language.
