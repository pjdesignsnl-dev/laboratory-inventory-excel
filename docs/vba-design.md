# VBA design — Laboratory Inventory for Excel

**Status:** Design for the frozen application contract (D-021)
**Contract:** `schema/workbook-contract.yaml` `status: frozen`, `contract_version: 1.0.0`
**Architecture source commit:** `731b7a13f6287aa843f92ee2dd8cb48f5d8b1111`
**Workbook SHA-256:** `C3D27FE82840833459674690DA250EC1500E99CC9337E52F4A7C4FAF81ED787A`
**Excel version:** 16.0.20228.20190 (x64)
**Date:** 2026-08-17

---

## 1. Engineering principles (Phase B mapping)

1. **Business logic separation** — worksheet events only capture input and
   dispatch to service routines in standard modules. Validation, transition
   rules, transaction construction, state mutation, ID generation, and rollback
   live in modules/classes, never in worksheet event procedures.
2. **Runtime contract validation** — `modWorkbookContract` validates at startup
   (and before each mutation) that worksheets, ListObjects, columns, named
   ranges, controlled statuses and transaction types exactly match the frozen
   contract. On drift: fail closed with a clear diagnostic; no partial
   operation.
3. **Source control** — every module is exported as text to `vba/`; the
   embedded VBA and `vba/` text are verified to correspond (export-compare
   script in Phase E).
4. **No fragile references** — `ThisWorkbook`, `Worksheets("...")`,
   `ListObjects("...")`, named ranges, header/index resolution only. No
   `ActiveWorkbook`/`ActiveSheet`/`Selection`/`Select`/`Activate` except a
   narrowly documented UI-focus operation on Scan.
5. **Application state safety** — any routine that changes
   `EnableEvents`/`ScreenUpdating`/`Calculation`/`DisplayAlerts`/`StatusBar`/
   protection restores prior state in a guaranteed cleanup (try/finally
   pattern), including error paths. Never assume default application state.
6. **Re-entrancy / duplicate-scan guard** — a module-level `m_busy` flag plus a
   pending-operation token prevents repeated Enter events, event re-entry, and
   rapid duplicate scans of the *same pending operation*. Legitimate later
   scans of the same container are allowed (the guard is cleared when the
   operation commits or aborts).
7. **Atomic mutation strategy** — explicit two-phase commit:
   *Phase 1 (prepare):* read old container state; validate all required
   fields; construct the transaction row in memory; verify destination
   rows/IDs are writable; capture old application state.
   *Phase 2 (commit):* append transaction row; update container row; verify
   post-conditions; recalc.
   On any write failure during commit: roll back every mutated container
   field, remove the transaction row **only if it belongs to the current
   uncommitted operation and has not become an accepted audit event**, restore
   formulas/protection/application state. Historical committed transactions
   are never deleted; compensation (`Adjustment`) is used where true rollback
   cannot be guaranteed.
8. **Post-condition validation** — after every successful mutation verify:
   unique ContainerID, unique Barcode, valid ProductID/LocationID, valid
   status, transaction exists exactly once, snapshot matches before/after,
   stock reconciles, no formula errors. Fail loudly otherwise.
9. **ID generation** — MAX over frozen helper columns + 1 (never row count);
   uniqueness verified immediately before commit. Receive N produces N unique
   ContainerIDs, N unique Barcodes, N Receive transactions.
10. **Time / user abstraction** — `modUtilities.GetNow()` and
    `modUtilities.GetOperator()` centralize clock and Windows-username capture
    (`Environ("USERNAME")`); no scattered `Now`/`Date`.
11. **Error classification** — blocking validation error / confirmation
    warning / informational message / unexpected runtime error. Operator-facing
    messages are clear; raw internals only in a diagnostics log.
12. **Backup safety** — a required backup that fails is treated as a failed
    operation (no silent continue). Timestamped unique filenames; never
    overwrite the only known-good backup.
13. **Performance** — bulk reads/writes via arrays and dictionaries; use
    ListObjects as structural boundary; design for thousands of Containers and
    tens of thousands of Transactions without per-cell round trips.
14. **Macro security** — document signed VBA / trusted publisher where
    practical, controlled trusted location only if deliberately chosen,
    Mark-of-the-Web considerations. Never instruct the user to globally enable
    all macros.

---

## 2. Frozen contract members referenced by VBA

### 2.1 Worksheets (exact names, frozen)
`Dashboard`, `Scan`, `Receiving`, `Products`, `Containers`, `Transactions`,
`Suppliers`, `Locations`, `Settings`

### 2.2 Tables (exact names + columns; order = frozen)

**tblProducts** (Products):
ProductID, ProductName, ProductType, Category, Manufacturer,
ManufacturerCatalogueNumber, CASNumber, Concentration, Grade,
StandardContainerDescription, SupplierID, StorageRequirements,
HazardClassification, SDSReference, MinimumContainerStock,
TargetContainerStock, ReorderQuantity, Active, Notes,
HelperAvailableStock (calc), HelperStockClass (calc)

**tblContainers** (Containers):
ContainerID, Barcode, ProductID, BatchLotNumber, ExpiryDate, RetestDate,
DateReceived, StorageLocationID, Status, OpenedDate, DisposalDate,
DisposalReason, Notes, HelperContainerNum (calc), HelperBarcodeNum (calc)

**tblTransactions** (Transactions, append-only):
TransactionID, Timestamp, Operator, Barcode, ContainerID, ProductID,
ProductName, TransactionType, PreviousStatus, NewStatus, PreviousLocation,
NewLocation, BatchLotNumber, Reason, Reference, Notes

**tblSuppliers / tblLocations / tblSettings / tblStatusList /
tblTransactionTypeList / tblExpiryClassList / tblScanResults (staging) /
tblReceiveStaging (staging)** — exact columns per the frozen YAML.

### 2.3 Named ranges (exact)
`rngScanInput` (Scan!D7), `rngScanStatusMessage` (Scan!D9),
`rngScanResultCard` (Scan!D4:K34); Receiving: `rngReceiveProductID` (D7),
`rngReceiveNextContainerID` (D9), `rngReceiveNextBarcode` (D10),
`rngReceiveLot` (D12), `rngReceiveExpiry` (D13), `rngReceiveRetest` (D14),
`rngReceiveLocation` (D15), `rngReceiveQuantity` (D16),
`rngReceiveStatusMessage` (D19); settings names `ExpiryWarningDays30/60/90`,
`DefaultLocationID`, `DefaultStatusNewContainers`; list names
`lstProductType`, `lstCategory`, `lstLocationType`, `lstDisposalReason`,
`lstTransactionReason`, `lstBool`; column-list names
`lstProductsProductID`, `lstProductsProductName`, `lstStatusList`,
`lstTransactionTypeList`, `lstExpiryClassList`, `lstSupplierIDs`,
`lstLocationIDs`.

### 2.4 Controlled values
Statuses (6): `Available`, `InUse`, `Expired`, `Damaged`, `Disposed`,
`Missing`.
Transaction types (9): `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`,
`MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment`.
Product types: `Consumable`, `Chemical`, `Reagent`.
Categories: `General`, `Pipette Tips`, `Tubes`, `Solvent`, `Reagent`,
`Consumable`.
Location types: `Cabinet`, `Shelf`, `Fridge`, `Freezer`, `Room`, `Other`.
Disposal reasons: `Used Up`, `Expired`, `Damaged`, `Missing`, `Other`.
Transaction reasons: `Used Up`, `Expired`, `Damaged`, `Missing`,
`Correction`, `Other`.
Booleans: `TRUE`, `FALSE`.

### 2.5 ID / barcode formats
ContainerID `C######`; ProductID `P######`; TransactionID `T########`;
SupplierID `S######`; StorageLocationID `LOC####`; Barcode `\d{7}` text.

### 2.6 Stock / expiry semantics (D-018)
Usable available = `Status="Available" AND (ExpiryDate blank OR ExpiryDate
>= TODAY())`. Expired-by-date `Available` containers are excluded from usable
stock and blocked from TakeOpen; Status is never silently mutated by the clock.

---

## 3. Module inventory and responsibilities

### 3.1 Standard modules

| Module | Responsibility | Key public members |
|---|---|---|
| `modConstants` | All contract constants: sheet/table/column names (exact), statuses, transaction types, ID formats, named ranges, expiry bands, message classes, misc literals. | `Const`/`Enum` declarations only; `cs*`/`tc*`/`st*` prefixes |
| `modWorkbookContract` | Runtime contract validator. Checks presence + exact names of worksheets, ListObjects, required columns, critical named ranges, controlled statuses, transaction types. | `ContractValidate() As Boolean`, `ContractDiagnostics() As String`, `FailClosed(msg)` |
| `modUtilities` | Time/user abstraction, string/date helpers, application-state save/restore, array helpers. | `GetNow()`, `GetOperator()`, `WithAppState(proc)`, `IsText7DigitBarcode()`, `NormalizeBarcode()` |
| `modBarcodeLookup` | Resolve barcode → Container row / staging display; duplicate detection. | `LookupByBarcode(bc, ByRef cid, ByRef status, ...) As LookupResult`, `FindContainerRow(bc) As Long`, `BarcodeCount(bc) As Long` |
| `modValidation` | All validation rules: transition matrix, required fields, FK checks, expiry-by-date blocks, duplicate/unknown handling. | `ValidateTransition(cur, txnType, ByRef message, ByRef class) As Boolean`, `ValidateRequiredReceive(...)`, `IsExpiredByDate(row)` |
| `modTransactions` | Transaction ID generation, snapshot construction, atomic append, post-condition verification. | `NextTransactionID() As String`, `BuildSnapshot(txnType, containerRow, ...) As Variant`, `AppendTransaction(snapshot) As String`, `VerifyAppended(tid) As Boolean` |
| `modContainers` | Container state mutation + rollback; ID generation for containers; uniqueness. | `NextContainerID() As String`, `NextBarcode() As String`, `ApplyStateChange(containerRow, newStatus, newLocation, ...)`, `RollbackContainerState(saved)` |
| `modReceiving` | Receive-one and Receive-N workflows; staging population; batch validation. | `ReceiveOne(productID, lot, expiry, retest, location, qty?) As String`, `ReceiveN(productID, lot, expiry, retest, location, n) As String()` |
| `modScanInterface` | Scan input handling (dispatch from worksheet event), transaction picker, confirm, reset/focus, duplicate-scan guard, message display. | `HandleScannedBarcode(bc)`, `CommitPending()`, `ResetScanField()`, `FocusScanInput()`, `m_busy` flag |
| `modBackup` | Backup/recovery: timestamped unique copies, restore instructions, failure handling. | `CreateBackup(Optional note) As String`, `BackupRequired() As Boolean`, `RestoreFromBackup(path)` |
| `modErrorHandling` | Centralized error classification, logging, operator messages, diagnostics. | `HandleError(errObj, context)`, `ClassifyError(...)`, `LogError(...)` |
| `modTestHooks` | Expose internal procedures for VBA tests (compile-only or via a test sheet). | `Test_RunAll()`, `Test_...` per workflow, `Test_ResetFixtures()` |

### 3.2 Document modules

| Module | Responsibility |
|---|---|
| `ThisWorkbook` | `Workbook_Open` → contract validation (fail closed with diagnostic), backup-on-open policy hook, initialize state. `Workbook_BeforeClose` → final state checks, optional backup. |
| `Scan` (worksheet module) | `Worksheet_Change` on `rngScanInput` → dispatch to `modScanInterface.HandleScannedBarcode`; nothing else. |
| `Receiving` (worksheet module) | `Worksheet_Change` on `rngReceiveProductID`/lot/expiry/location/quantity → dispatch to `modReceiving.UpdateStaging`; `Worksheet_SelectionChange` only if a receive button is implemented (UI-focus only). |

Events are **dispatch-only**; all business rules live in the standard modules.

---

## 4. Data flow and atomicity boundary

### 4.1 Scan → mutation pipeline
1. Scanner/typing writes digits + Enter to `rngScanInput` (Scan!D7).
2. `Scan.Worksheet_Change` fires → guard `m_busy`; dispatch to
   `modScanInterface.HandleScannedBarcode`.
3. `modBarcodeLookup` resolves barcode → staging (`tblScanResults`) displays
   container/product/status/location/expiry-class.
4. `modValidation` computes allowed next actions and blocking conditions from
   the frozen matrix + D-018.
5. Operator selects a transaction type (staging cell / picker).
6. `modScanInterface.CommitPending` runs the **atomic prepare/commit**:
   - prepare: build snapshot, validate, capture old state;
   - commit: append transaction, mutate container;
   - post-conditions; recalc; reset scan field; focus back to `rngScanInput`.

### 4.2 Atomicity boundary (documented)
- **Boundary:** one transaction row + one container row per operation.
- **Prepare** must complete before any write; if prepare fails, nothing is
  written.
- **Commit** is ordered: append transaction row **first**, then mutate
  container. If container mutation fails after the transaction row was
  appended, the transaction row is removed **only if** (a) it was created by
  this operation and (b) it has not been externally accepted (no other code
  referenced it); then container rollback restores old fields.
- If rollback itself fails, a compensation `Adjustment` transaction is appended
  recording the intended state; the UI reports the failure loudly.
- Historical committed transactions are never deleted.

---

## 5. Runtime contract validator (modWorkbookContract)

`ContractValidate()` checks, in order (fail fast):
1. All 9 worksheets exist (exact names).
2. Each expected ListObject exists on its worksheet (exact names).
3. Each required column exists in the expected position/order (header text ==
   column name, frozen).
4. Critical named ranges exist and resolve to the frozen addresses.
5. `tblStatusList` contains exactly the 6 frozen statuses.
6. `tblTransactionTypeList` contains exactly the 9 frozen transaction types.
7. `tblSettings` contains the expiry-band keys and default location/status.
8. Barcode column cells are text (leading zeros preserved) — data-level spot
   check on fixtures; full check on demand.

On any failure: `FailClosed(msg)` shows a blocking message box with the
diagnostic (sheet/table/column/name + expected vs actual), disables the scan
input, and prevents all mutation entry points. The workbook cannot be used for
inventory operations until the drift is corrected.

---

## 6. State-transition implementation

`modValidation.ValidateTransition(curStatus As String, txnType As String,
ByRef message As String, ByRef msgClass As MsgClass) As Boolean` implements the
frozen matrix exactly:

- `Available` (valid date): TakeOpen ALLOW; Transfer ALLOW; Dispose CONFIRM;
  MarkExpired/MarkDamaged/MarkMissing CONFIRM; Receive/Return/Adjustment per
  matrix.
- `Available` but expired-by-date (D-018): TakeOpen **BLOCK**; Transfer ALLOW
  (relocation); Dispose ALLOW; MarkExpired ALLOW (recommended);
  MarkDamaged/MarkMissing CONFIRM.
- `InUse`: Return ALLOW; Dispose/Mark* CONFIRM; TakeOpen BLOCK; Transfer BLOCK.
- `Expired`: Dispose ALLOW; Transfer ALLOW; others BLOCK/N/A per matrix.
- `Damaged`: Dispose ALLOW; Transfer ALLOW.
- `Disposed`: BLOCK everything except Adjustment CONFIRM.
- `Missing`: Dispose CONFIRM (resolve); Adjustment CONFIRM (found).
- `(none)` new container: Receive ALLOW.

Message classes: `mcBlocking`, `mcConfirm`, `mcInfo`.

---

## 7. Module-level procedures and test plan

For every module, the sections below record: exact worksheets/tables/columns
touched, named ranges, statuses/transaction types relied on, dependencies,
trigger, validation rules, error behaviour, rollback behaviour, tests.

(This section is completed in the implementation phase; see each stage's
evidence file under `evidence/vba/`.)

---

## 8. Stage plan (Phase D mapping)

| Stage | Module(s) | Deliverable | Test |
|---|---|---|---|
| 1 | modConstants, modWorkbookContract | Constants + runtime validator | compile; validator passes on frozen workbook; fails on drifted copy |
| 2 | modBarcodeLookup | barcode → container resolution | known/unknown/duplicate |
| 3 | modValidation | transition + D-018 validation | every allowed/blocked/confirm transition |
| 4 | modTransactions (ID + snapshot) | TransactionID gen + immutable snapshot | uniqueness, snapshot fields |
| 5 | modTransactions (atomic append) | atomic append + post-condition | append-once; failure injection |
| 6 | modContainers | state mutation + rollback | each transition; rollback on injected failure |
| 7 | modReceiving (ReceiveOne) | receive one | ID/barcode/txn correct |
| 8 | modReceiving (ReceiveN) | receive N identical | N unique IDs/barcodes/txns |
| 9 | modScanInterface | Enter event / reset / focus / guard | scan flow, duplicate-scan guard, focus |
| 10 | modUtilities.GetOperator | Windows username capture | returns non-empty |
| 11 | modBackup | backup/recovery | unique names, failure handling, restore |
| 12 | modErrorHandling | error classification + diagnostics | each class; log written |
| 13 | Code 128 | label preparation/printing support (printer-independent) | renders Code128 pattern/string for a barcode |

---

## 9. Macro security / deployment note (frozen architecture unaffected)

VBA sources are exported to `vba/` and embedded in the `.xlsm`. Deployment
(macro security settings, trusted location, code signing) is documented in the
final release; the frozen application architecture is independent of the
unresolved storage/deployment option (D-013 Proposed).
