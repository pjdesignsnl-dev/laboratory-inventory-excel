# Workbook architecture — Laboratory Inventory for Excel v0.1

**Status:** DRAFT — not approved for VBA
**Date:** 2026-08-17
**Contract reference:** `schema/workbook-contract.yaml` (status `draft`)
**Review scope:** Stages 2–3 of `docs/INITIAL_TASK.md`

---

## 1. Design summary

`LabInventory_v0.1.xlsx` is a macro-free Excel application that tracks physical laboratory containers by unique internal barcodes. It implements:

- **Nine worksheets** with a clear editor/viewer split.
- **Nine Excel Tables** with exact stable names (`tbl` prefix) and PascalCase columns.
- **Formula-driven stock, expiry, reorder, and dashboard logic** — no manually editable stock balances.
- **A polished Scan interface** that is fully functional for lookup and validation display without VBA.
- **A Receiving interface foundation** with guided fields and computed next IDs.
- **Validation, conditional formatting, and a protection design** that turns accidental damage into the exception, not the norm.
- **Synthetic worked examples** across five product families with complete container histories.

All formulas are Excel 2021/2024-compatible; Microsoft 365-only alternatives are labeled. **No VBA, no Office Scripts, no other automation layer exists anywhere in this phase.**

---

## 2. Workbook topology

| Worksheet | Purpose | Editors | Viewers | Excel Tables | Protected areas |
|---|---|---|---|---|---|
| `Dashboard` | Operational KPIs: totals, low/out-of-stock, expiry, recent activity | None (all formula output) | Everyone | `tblDashboard` (one-row stats) | All cells |
| `Scan` | Barcode lookup, container/product detail, validation, allowed next actions; transaction staging | Interim manual workflow + (future) VBA | Everyone | `tblScanResults` (lookup staging) | Lookup/detail area; entry cells remain editable |
| `Receiving` | Guided receiving interface foundation: product picker, ID generation, data-entry block, batch size | Store operator | Everyone | `tblReceiveStaging` (staging) | Instructions/IDs area; entry cells editable |
| `Products` | Product catalogue | Store operator | Everyone | `tblProducts` | Header + formula columns |
| `Containers` | Container master (barcode, lot, expiry, location, status) | Store operator | Everyone | `tblContainers` | Header + formula columns |
| `Transactions` | Append-only transaction history | Store operator (append only) | Everyone | `tblTransactions` | Header; whole-table protection guidance |
| `Suppliers` | Supplier catalogue | Store operator | Everyone | `tblSuppliers` | Header |
| `Locations` | Storage locations | Store operator | Everyone | `tblLocations` | Header |
| `Settings` | Controlled lists + configurable constants (30/60/90-day bands) | Administrator | Everyone | `tblSettings`, `tblStatusList`, `tblTransactionTypeList`, `tblExpiryClassList` | All cells (read-mostly) |

### 2.1 Sheet order and navigation
`Dashboard` → `Scan` → `Receiving` → `Products` → `Containers` → `Transactions` → `Suppliers` → `Locations` → `Settings`. A `Navigation` named range per sheet (tab color coding: dashboard navy, scan green, receiving teal, data sheets blue, lists gray) supports consistent location.

---

## 3. Table definitions

Table naming: `tbl` + PascalCase entity. Column naming: PascalCase, stable. All barcode/ID columns are **text**; all amounts are integers; all dates are real dates with `yyyy-mm-dd` display.

### 3.1 `tblProducts` (sheet `Products`)

Primary key: `ProductID` (unique, nonblank).

| Column | Data type | Mandatory | Default | Validation / notes |
|---|---|---|---|---|
| ProductID | Text (PK) | Yes | — | Pattern `P######`; unique |
| ProductName | Text | Yes | — | Nonblank |
| ProductType | Text (list) | Yes | Consumable | List: `Consumable`, `Chemical`, `Reagent` |
| Category | Text (list) | Yes | General | List: `General`, `Pipette Tips`, `Tubes`, `Solvent`, `Reagent`, `Consumable` |
| Manufacturer | Text | Yes | — | Nonblank |
| ManufacturerCatalogueNumber | Text | Yes | — | Nonblank |
| CASNumber | Text | No | — | Optional; chemicals only |
| Concentration | Text | No | — | Optional; free text (e.g., "96%") |
| Grade | Text | No | — | Optional (e.g., "Analytical") |
| StandardContainerDescription | Text | No | — | Optional |
| SupplierID | Text (FK) | Yes | S000001 | Must exist in `tblSuppliers` |
| StorageRequirements | Text | No | — | Optional |
| HazardClassification | Text | No | — | Optional |
| SDSReference | Text | No | — | Optional |
| MinimumContainerStock | Integer | Yes | 0 | >= 0 |
| TargetContainerStock | Integer | Yes | 1 | >= 0 |
| ReorderQuantity | Integer | Yes | 1 | >= 0 |
| Active | Boolean | Yes | TRUE | `TRUE`/`FALSE` list |
| Notes | Text | No | — | Optional |

**Formula columns (protected):** none in the table; stock/reorder computed on Dashboard and Products via COUNTIFS against `tblContainers` (not stored — integrity).

### 3.2 `tblContainers` (sheet `Containers`)

Primary key: `ContainerID` (unique, nonblank). Alternate key: `Barcode` (unique, nonblank).

| Column | Data type | Mandatory | Default | Validation / notes |
|---|---|---|---|---|
| ContainerID | Text (PK) | Yes | — | Pattern `C######`; unique; immutable |
| Barcode | Text (alternate key) | Yes | — | Pattern `\d{7}`; unique; stored as text |
| ProductID | Text (FK) | Yes | — | Must exist in `tblProducts` |
| BatchLotNumber | Text | Yes | — | Nonblank (synthetic: `LOT####`) |
| ExpiryDate | Date | No | — | Optional; >= DateReceived if present |
| RetestDate | Date | No | — | Optional advisory only |
| DateReceived | Date | Yes | — | <= TODAY (validation) |
| StorageLocationID | Text (FK) | Yes | LOC0001 | Must exist in `tblLocations` |
| Status | Text (list) | Yes | Available | List from `tblStatusList`: `Available`, `InUse`, `Reserved`, `Expired`, `Damaged`, `Disposed`, `Missing` |
| OpenedDate | Date | No | — | Set when opened (flag semantics) |
| DisposalDate | Date | No | — | Set on disposal/expiry-marking |
| DisposalReason | Text (list) | No | — | List: `Used Up`, `Expired`, `Damaged`, `Missing`, `Other` |
| Notes | Text | No | — | Optional |

**Formula columns (protected):** none stored; lookups (product name, expiry class, location name, stock position) are formula-driven on Dashboard/Scan/Products.

### 3.3 `tblTransactions` (sheet `Transactions`)

Primary key: `TransactionID` (unique, nonblank). **Append-only** — invariant 7.

| Column | Data type | Mandatory | Default | Validation / notes |
|---|---|---|---|---|
| TransactionID | Text (PK) | Yes | — | Pattern `T########`; unique |
| Timestamp | Date+time | Yes | — | Real datetime; `yyyy-mm-dd hh:mm` |
| Operator | Text | No | — | Reserved; Windows username in VBA phase |
| Barcode | Text | Yes | — | References `tblContainers[Barcode]` (FK) |
| ContainerID | Text (FK) | Yes | — | References `tblContainers[ContainerID]` |
| ProductID | Text (FK) | Yes | — | References `tblProducts[ProductID]` |
| ProductName | Text | Yes | — | Snapshot at transaction time |
| TransactionType | Text (list) | Yes | — | `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment` |
| PreviousStatus | Text | Yes | — | From `tblStatusList`; `(none)` for Receive |
| NewStatus | Text | Yes | — | From `tblStatusList` |
| PreviousLocation | Text | Yes | — | `(none)` for Receive |
| NewLocation | Text | Yes | — | Target location |
| BatchLotNumber | Text | Yes | — | Snapshot |
| Reason | Text (list) | No | — | Optional: `Used Up`, `Expired`, `Damaged`, `Missing`, `Correction`, `Other` |
| Reference | Text | No | — | Optional (e.g., purchase order) |
| Notes | Text | No | — | Optional |

### 3.4 `tblSuppliers` (sheet `Suppliers`)

Primary key: `SupplierID`.

| Column | Data type | Mandatory | Validation / notes |
|---|---|---|---|
| SupplierID | Text (PK) | Yes | Pattern `S######`; unique |
| SupplierName | Text | Yes | Nonblank |
| ContactName | Text | No | |
| Email | Text | No | Contains `@` when present |
| Phone | Text | No | |
| Website | Text | No | |
| Address | Text | No | |
| Notes | Text | No | |

### 3.5 `tblLocations` (sheet `Locations`)

Primary key: `StorageLocationID`.

| Column | Data type | Mandatory | Validation / notes |
|---|---|---|---|
| StorageLocationID | Text (PK) | Yes | Pattern `LOC####`; unique |
| LocationName | Text | Yes | Nonblank |
| LocationType | Text (list) | Yes | `Cabinet`, `Shelf`, `Fridge`, `Freezer`, `Room`, `Other` |
| Description | Text | No | |
| Active | Boolean | Yes | `TRUE`/`FALSE` |

### 3.6 `tblSettings` (sheet `Settings`)

Key/value constants. No primary key column; keys unique.

| Column | Data type | Mandatory | Notes |
|---|---|---|---|
| SettingKey | Text (PK) | Yes | Unique |
| SettingValue | Text | Yes | Stored as text; numeric keys parsed by formulas |
| Description | Text | No | |

Keys: `ExpiryWarningDays30`, `ExpiryWarningDays60`, `ExpiryWarningDays90`, `DefaultLocationID`, `DefaultStatusNewContainers`, `ScannerEnterSuffix` (informational), `WorkbookVersion`.

### 3.7 List tables (sheet `Settings`)

- `tblStatusList` — `StatusValue` (PK, unique): the 7 statuses.
- `tblTransactionTypeList` — `TransactionTypeValue` (PK, unique): the 9 types.
- `tblExpiryClassList` — `ExpiryClassValue` (PK, unique): `Expired`, `ExpiringSoon`, `Valid`, `NoExpiry`, `Invalid` plus a `Label` column for display.

### 3.8 Staging tables (non-authoritative, UI-only)

- `tblScanResults` (sheet `Scan`) — holds the resolved container/product lookup for the entered barcode: `Barcode`, `ContainerID`, `ProductID`, `ProductName`, `BatchLotNumber`, `ExpiryDate`, `StorageLocationID`, `LocationName`, `Status`, `OpenedDate`, `ExpiryClass`, `DuplicateFlag`, `LookupState` (e.g., `FOUND`, `UNKNOWN`, `DUPLICATE`, `EMPTY`). Populated by formulas; **never a source of truth**.
- `tblReceiveStaging` (sheet `Receiving`) — holds the current receive entry: `ProductID`, `ProductName`, `NextContainerID`, `NextBarcode`, `BatchLotNumber`, `ExpiryDate`, `RetestDate`, `StorageLocationID`, `Status`, `Quantity`, `ValidationMessage`, `ReadinessState`. Populated by formulas from entry cells.

---

## 4. Relationships

```
tblProducts       1 ──── * tblContainers        (ProductID)
tblSuppliers      1 ──── * tblProducts          (SupplierID)
tblLocations      1 ──── * tblContainers        (StorageLocationID)
tblContainers     1 ──── * tblTransactions      (ContainerID)
tblProducts       1 ──── * tblTransactions      (ProductID)
tblTransactions   * ──── 1 tblContainers        (Barcode)
```

- Referential integrity is enforced by validation lists (Supplier/Location/Product) and by COUNTIFS-based "orphan" checks on Dashboard/Scan; hard foreign keys are a VBA-phase enforcement target.
- `tblTransactions` is the audit spine; `Barcode` and `ContainerID` are both carried so a barcode search reconstructs complete history without joining (Req §6).
- `ProductName`, `BatchLotNumber` are snapshotted on Transactions (history must survive product edits — Req §9).

---

## 5. Controlled lists and constants

| List | Source | Values |
|---|---|---|
| ProductType | `tblStatusList`? No — inline list | `Consumable`, `Chemical`, `Reagent` |
| Category | inline list | `General`, `Pipette Tips`, `Tubes`, `Solvent`, `Reagent`, `Consumable` |
| ContainerStatus | `tblStatusList` | `Available`, `InUse`, `Reserved`, `Expired`, `Damaged`, `Disposed`, `Missing` |
| TransactionType | `tblTransactionTypeList` | `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment` |
| ExpiryClass | `tblExpiryClassList` | `Expired`, `ExpiringSoon`, `Valid`, `NoExpiry`, `Invalid` |
| LocationType | inline list | `Cabinet`, `Shelf`, `Fridge`, `Freezer`, `Room`, `Other` |
| DisposalReason | inline list | `Used Up`, `Expired`, `Damaged`, `Missing`, `Other` |
| Reason (Transactions) | inline list | `Used Up`, `Expired`, `Damaged`, `Missing`, `Correction`, `Other` |
| Active / flags | inline list | `TRUE`, `FALSE` |

Inline lists are defined as named ranges (`lstProductType`, `lstCategory`, `lstLocationType`, `lstDisposalReason`, `lstTransactionReason`, `lstBool`).

**Constants (`tblSettings`):** expiry bands 30/60/90 days; default location; default new-container status; workbook version.

---

## 6. Product vs batch vs container data placement

| Attribute | Level | Rationale |
|---|---|---|
| Catalogue identity, manufacturer, CAS, concentration, grade, hazard, SDS, thresholds | Product | One per catalogue item |
| Supplier reference | Product | Supplier is a product attribute in v1 |
| Batch/lot, expiry/retest dates, received date, location, status, opened/disposal | Container | Physical-instance attributes |
| Transaction history | Transaction | Event-level; snapshots reference product and container |

No separate batch table in v1: a received lot is a Container attribute. Batch-level separation is a documented future option if the owner needs lot-level lot numbers/COA tracking (open question).

---

## 7. Barcode design

- **Format:** 7-digit numeric (`0000001` … `9999999`), zero-padded, stored as **text** to preserve leading zeroes (default assumption 7).
- **Uniqueness:** exactly one Container per Barcode (invariant 3); duplicate detection is formula-driven (`DuplicateFlag` on Scan and Containers via COUNTIF > 1) and will be hard-enforced by VBA.
- **Uniqueness of ContainerID:** immutable `C######`, distinct from barcode (default assumption 8).
- **Label symbology:** Code 128 (recommended), internal label per container; supplier/manufacturer barcodes cannot identify identical containers (Req §17).
- **Scanner config:** USB keyboard-wedge, **Enter suffix**, output barcode digits followed by Enter; scan field auto-clears by VBA later. Documented in `docs/` deployment notes.
- **Scan semantics:** a scan of an empty field is ignored (lookup state `EMPTY`); a scan of an unknown barcode shows `UNKNOWN`; a duplicate shows `DUPLICATE` (blocking in VBA phase).

---

## 8. Status and transaction model

See `docs/status-transition-matrix.md` for the full matrix. Summary:

- **Statuses:** `Available`, `InUse`, `Reserved`, `Expired`, `Damaged`, `Disposed`, `Missing`.
- **Opened is a flag** (`OpenedDate`), not a status. A returned opened container may be `Available` again while retaining opened history (default assumption 12).
- **Transactions:** `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment`.
- **Append-only:** history is never overwritten; corrections are compensating `Adjustment` transactions (invariant 7).
- **Atomicity:** VBA-phase requirement; the interim manual workflow documents exact row-append order (Container update then Transaction append, both before any next action).

---

## 9. Stock, expiry, and reorder logic

All figures below are **formula-derived**, never stored as editable numbers (default assumption 10; Req §12).

### 9.1 Available stock
```
AvailableCount = COUNTIFS(tblContainers[ProductID], <pid>, tblContainers[Status], "Available")
```
Statuses qualifying as available stock: **`Available` only** (decision D-006). `InUse`/`Reserved`/terminal states are excluded.

### 9.2 Reorder classification (per product)
```
if AvailableCount = 0                       -> "OutOfStock"
elif AvailableCount < MinimumContainerStock -> "Reorder"
elif AvailableCount < TargetContainerStock  -> "Low"
else                                        -> "OK"
```

### 9.3 Expiry classification (per container)
```
if ExpiryDate blank                  -> "NoExpiry"
if ExpiryDate < TODAY()              -> "Expired"
if ExpiryDate <= TODAY()+30          -> "ExpiringSoon"  (band 30)
if ExpiryDate <= TODAY()+60          -> "ExpiringSoon"  (band 60)
if ExpiryDate <= TODAY()+90          -> "ExpiringSoon"  (band 90)
else                                 -> "Valid"
```
Bands are read from `tblSettings` (`ExpiryWarningDays30/60/90`); a container is assigned the **tightest applicable band**. `ExpiringSoon` is subdivided in the Dashboard into 30/60/90 buckets via COUNTIFS.

### 9.4 Dashboard statistics
- Total active products: `COUNTIFS(tblProducts[Active], TRUE)`
- Total available containers: `COUNTIF(tblContainers[Status], "Available")`
- Low-stock products: `COUNTIFS` over a per-product helper column
- Out-of-stock products: same
- Expired containers: `COUNTIF(tblContainers[ExpiryDate],"<"&TODAY())` (status-independent for the alarm) plus status-aware variants
- Expiring soon (30/60/90): COUNTIFS on expiry bands
- Recently received/taken (last 14 days): COUNTIFS on `tblTransactions[Timestamp]`
- Frequently used: `COUNTIF` on `tblTransactions[ProductID]` (top by count)
- Inventory by category/location: COUNTIFS joins via helper columns

---

## 10. Scan workflow

The Scan sheet implements the lookup-and-validate half of Req §4 end-to-end **without VBA**:

1. User scans (or types) a barcode into `rngScanInput` (B7 on Scan).
2. `tblScanResults` formulas resolve Container/Product/Location/Status/ExpiryClass/DuplicateFlag.
3. Detail cards display Product, Container, Status, Location, Expiry class, Opened status.
4. Validation panel computes `AllowedNextActions` from the current status (mirror of the transition matrix) and flags `UNKNOWN`/`DUPLICATE`/`DISPOSED`/`EXPIRED`/`MISSING` conditions as blocking warnings.
5. In the interim manual mode, the operator performs the transaction on the Containers and Transactions sheets (validation/CF guides correctness); in the VBA phase, a transaction picker + confirm writes atomically.
6. `rngScanInput` is reset and re-focused by VBA later (VBA-phase feature; documented, not implemented).

**Exact Scan layout (fixed for contract):** A1 title; B4 scan label; B7 scan input; D4–K12 container detail card; D14–K24 product detail card; D26–K34 validation panel; B37 instructions. Named ranges: `rngScanInput` (B7), `rngScanStatusMessage` (B9), `rngScanResultCard` (D4:K34). **These exact addresses are contract-bound.**

---

## 11. Receiving workflow

The Receiving sheet is the foundation (Req §17) with exact staging cells:

1. Select Product (`rngReceiveProductID`, B7) → product details auto-fill.
2. `rngReceiveNextContainerID` (B9) and `rngReceiveNextBarcode` (B10) are computed from the max existing ID/barcode + 1 (formula, text-preserving).
3. Enter lot (`rngReceiveLot`, B12), expiry (`rngReceiveExpiry`, B13), retest (B14), location (`rngReceiveLocation`, B15).
4. Enter quantity (`rngReceiveQuantity`, B16, default 1).
5. Validation panel (`rngReceiveStatusMessage`, B19) checks completeness/duplicates.
6. Interim manual mode: append one `tblContainers` row and one `Receive` `tblTransactions` row per container; VBA phase will execute the atomic multi-row append.

---

## 12. Dashboard and reporting

One-row `tblDashboard` KPI row plus a series of COUNTIFS-driven stat blocks (totals, low/out-of-stock table, expiry band table, recent activity, category/location breakdown, frequently-used top-5). Every number traces to a Table via COUNTIFS — the reconciliation test (C-004) proves Dashboard == source.

---

## 13. Protection and manual-edit controls

- **Locked by default:** all formula cells, headers, instruction blocks, and ID/read-only areas.
- **Editable (unlocked):** data-entry cells only (entry blocks on Scan/Receiving; Table body columns on data sheets — Tables normally allow body edits; protection applies to formula columns and headers via sheet protection + locked style).
- **Sheet protection** with `SelectLockedCells` allowed but edits blocked; protection password **empty** (deterrent, not security — Req §24) with the caveat documented.
- **Workbook structure protection** (prevent sheet rename/delete) with empty password; documented.
- **Conditional formatting:** red for `Expired`, amber for `ExpiringSoon`, green for `Available`/`Valid`, gray for terminal states; duplicate barcode cells highlighted.
- **Append-only enforcement:** no delete shortcuts on Transactions in instructions; VBA-phase hard enforcement; versioned backups recommended.

---

## 14. Read-only and deployment architecture

### 14.1 Options comparison (Req §20)

| Option | Pros | Cons | Verdict |
|---|---|---|---|
| **A. One master `.xlsm` on network share** | Single source of truth; one writer PC enforced by file open; simple backup; Excel-native | Network latency; file lock on open; corruption risk on interrupted writes; no offline editing | **Recommended** for one writer + multiple readers |
| B. Master + read-only reporting workbook | Readers never touch master; lighter locks; good audit distribution | Two artifacts to maintain; read-only copy staleness | Adopt **after v1.0.0** as `LabInventory_ReadOnly_v1.0.0.xlsx` |
| C. Microsoft 365/SharePoint/OneDrive | Co-authoring, versioning, web access | Co-authoring conflicts with single-writer design; web editing of `.xlsm` limited; sync failure modes; VBA not supported in web | Not recommended for the writer PC; possible read-only publishing later |
| D. Google Drive Office-file sync | Cheap sync, versioning | Conversion risk; file conflicts; not native Excel; VBA unsupported; violates "never convert to native Google Sheets" | Not recommended |

### 14.2 Recommendation
**Option A primary**: master workbook on a private network share; the dedicated laboratory PC is the only inventory writer; other PCs receive read-only copies (Option B pattern) after stabilization. Backups: daily copy to second location + weekly archive, retention 4 weeks, with a documented restore drill (VBA-phase backup/recovery module will automate; manual copies are the v0.1 baseline).

---

## 15. Compatibility

- All formulas are Excel 2021/2024-compatible (COUNTIFS/COUNTIF/IF/AND/OR/TODAY/EDATE/TEXT/MAX/SUMPRODUCT); **no** `FILTER`, `XLOOKUP`, `SORT`, `TEXTJOIN`, `LET`, or dynamic arrays in committed cells.
- Microsoft 365-only alternatives are labeled in comments/side notes (e.g., `FILTER` for the top-5 list) but not used in live formulas.
- Barcodes/IDs stored as text; dates as real dates; booleans as real booleans.
- Table structured references throughout; named ranges for Scan/Receiving interface cells.

---

## 16. Trade-offs, risks, and deferred scope

| Trade-off / risk | Description | Mitigation |
|---|---|---|
| Manual interim mode (no VBA) | Transaction appends are manual; atomicity depends on operator discipline | Guard rails: validation, CF, documented order, protection; VBA-phase contract requires atomic appends |
| No batch table in v1 | Lot tracked at container level only | Documented; optional batch table later |
| Statuses vs flags | `Opened` is a flag; keeps state count small | Documented in transition matrix |
| Protection ≠ security | Empty-password sheet protection is deterrent only | Documented; backups; VBA enforcement later |
| Desktop Excel unverified in this phase | No Excel on build machine | Independent formula evaluation + structural inspection + LibreOffice headless smoke test (if obtainable); owner desktop verification step defined |
| COUNTIFS performance at scale | Fine to ~100k rows | v0.1 fixture scale tiny; documented limit |
| Text barcodes vs numeric | Leading-zero preservation wins | Stored `@`; comparison in formulas is text-safe |

**Deferred to VBA phase (explicit):** atomic transaction append, container-state update, scan input reset/focus, duplicate-scan guard, receive batch helper, operator capture, backup/recovery automation, hard referential integrity.

---

## 17. Architecture review checklist

- [x] Every worksheet is defined (9 sheets + purpose + editors/viewers).
- [x] Every Excel Table has an exact stable name (9 Tables).
- [x] Every column has an exact stable name and data type.
- [x] Mandatory/optional rules are explicit.
- [x] Keys, uniqueness, and relationships are explicit.
- [x] Statuses and transactions have unambiguous semantics.
- [x] Every state transition is classified allowed/warning/blocked (see matrix).
- [x] Barcode uniqueness and duplicate-scan behavior are defined.
- [x] Stock is not manually editable.
- [x] Formula compatibility is documented (2021-compatible).
- [x] Scan and receiving interfaces have exact ranges/named ranges.
- [x] Protection and recovery limitations are documented.
- [x] Non-VBA tests are defined (`tests/test-plan.md`; executed in `tests/non-vba-results.md`).
