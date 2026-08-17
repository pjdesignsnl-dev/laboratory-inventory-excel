# Frozen workbook contract

**Status:** FROZEN FOR VBA (owner-approved architecture freeze, 2026-08-17)

This is the human-readable binding contract between the workbook and VBA.
**Frozen members may not change without the documented change-control process.**
The final storage/deployment option is deliberately **not** frozen (D-013
remains Proposed; see `docs/decisions.md` and the `deployment_status: proposed`
block in `schema/workbook-contract.yaml`).

## Workbook identity

- Product name: Laboratory Inventory for Excel
- Workbook version: v0.1
- Contract version: 1.0.0 (frozen)
- Target Excel versions: Microsoft 365 desktop, Excel 2021/2024 desktop
- Architecture approval reference: `docs/architecture.md`; `docs/decisions.md` D-001…D-021
- Architecture source commit: `731b7a13f6287aa843f92ee2dd8cb48f5d8b1111`
- Workbook SHA-256: `C3D27FE82840833459674690DA250EC1500E99CC9337E52F4A7C4FAF81ED787A`
- Excel version used for acceptance: Microsoft Excel 16.0.20228.20190 (x64)
- Excel runtime acceptance: **29/29 PASS** (`evidence/excel-runtime/excel-runtime-results.txt`)
- Supplementary regression: structural **51/51 PASS**; formula/business-rule **55/55 PASS**

## Worksheets (frozen, exact order)

1. `Dashboard`
2. `Scan`
3. `Receiving`
4. `Products`
5. `Containers`
6. `Transactions`
7. `Suppliers`
8. `Locations`
9. `Settings`

## Excel Tables and exact columns (frozen)

The authoritative machine-readable definition is `schema/workbook-contract.yaml`
(`status: frozen`). In summary:

- `tblProducts` (Products): ProductID, ProductName, ProductType, Category, Manufacturer, ManufacturerCatalogueNumber, CASNumber, Concentration, Grade, StandardContainerDescription, SupplierID, StorageRequirements, HazardClassification, SDSReference, MinimumContainerStock, TargetContainerStock, ReorderQuantity, Active, Notes, HelperAvailableStock (calculated), HelperStockClass (calculated)
- `tblContainers` (Containers): ContainerID, Barcode, ProductID, BatchLotNumber, ExpiryDate, RetestDate, DateReceived, StorageLocationID, Status, OpenedDate, DisposalDate, DisposalReason, Notes, HelperContainerNum (calculated), HelperBarcodeNum (calculated)
- `tblTransactions` (Transactions, append-only): TransactionID, Timestamp, Operator, Barcode, ContainerID, ProductID, ProductName, TransactionType, PreviousStatus, NewStatus, PreviousLocation, NewLocation, BatchLotNumber, Reason, Reference, Notes
- `tblSuppliers` (Suppliers): SupplierID, SupplierName, ContactName, Email, Phone, Website, Address, Notes
- `tblLocations` (Locations): StorageLocationID, LocationName, LocationType, Description, Active
- `tblSettings` (Settings): SettingKey, SettingValue, Description
- `tblStatusList` (Settings): StatusValue, Label
- `tblTransactionTypeList` (Settings): TransactionTypeValue, Label
- `tblExpiryClassList` (Settings): ExpiryClassValue, Label
- `tblScanResults` (Scan, staging, non-authoritative): Barcode, ContainerID, ProductID, ProductName, BatchLotNumber, ExpiryDate, StorageLocationID, LocationName, Status, OpenedDate, ExpiryClass, DuplicateFlag, LookupState
- `tblReceiveStaging` (Receiving, staging, non-authoritative): ProductID, ProductName, NextContainerID, NextBarcode, BatchLotNumber, ExpiryDate, RetestDate, StorageLocationID, Status, Quantity, ValidationMessage, ReadinessState

Column header text equals the column name exactly (so Excel structured
references resolve).

## Named ranges and exact cells (frozen)

Interface ranges (Scan/Receiving), list ranges, settings names, and column-list
names — exact addresses in `schema/workbook-contract.yaml`. Highlights:
`rngScanInput = Scan!D7`, `rngScanStatusMessage = Scan!D9`,
`rngScanResultCard = Scan!D4:K34`; `rngReceiveProductID = Receiving!D7`,
`rngReceiveNextContainerID = Receiving!D9`, `rngReceiveNextBarcode =
Receiving!D10`, `rngReceiveLot = Receiving!D12`, `rngReceiveExpiry =
Receiving!D13`, `rngReceiveRetest = Receiving!D14`, `rngReceiveLocation =
Receiving!D15`, `rngReceiveQuantity = Receiving!D16`, `rngReceiveStatusMessage
= Receiving!D19`; settings names `ExpiryWarningDays30/60/90`,
`DefaultLocationID`, `DefaultStatusNewContainers`.

## IDs and generation rules (frozen)

- ContainerID: `C######` (immutable, unique)
- ProductID: `P######`
- TransactionID: `T########`
- SupplierID: `S######`
- StorageLocationID: `LOC####`
- Barcode: `\d{7}` (7 digits, zero-padded, stored as text)
- Generation: MAX over the frozen helper columns (`HelperContainerNum`,
  `HelperBarcodeNum`) + 1, with uniqueness verified before commit. Row count is
  never used.

## Controlled status values (frozen, 6-value model D-016)

`Available`, `InUse`, `Expired`, `Damaged`, `Disposed`, `Missing`

## Controlled transaction types (frozen)

`Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`,
`MarkDamaged`, `MarkMissing`, `Adjustment`

Other controlled lists (product types, categories, location types, disposal
reasons, transaction reasons, booleans, expiry classes): exact values in
`schema/workbook-contract.yaml`.

## Allowed state transitions

See `docs/status-transition-matrix.md` (frozen; ALLOW/CONFIRM/BLOCK/N/A per
status and transaction). D-018 semantics: an `Available` container that is
expired by date is excluded from usable stock and blocked from TakeOpen without
silently mutating its stored Status.

## Formula contract

F-01…F-10 in `schema/workbook-contract.yaml`. All Excel 2021/2024-compatible.
Stock source of truth: `Status="Available" AND (ExpiryDate blank OR ExpiryDate
>= TODAY())`, implemented as COUNTIFS subtraction.

## Validation and protection contract

- List/custom data validation on entry columns; barcode 7-digit custom rule on
  `rngScanInput`.
- All 9 sheets protected (empty password — deterrent, not security, D-015).
- Workbook structure protected.
- `tblTransactions` append-only convention; corrections via `Adjustment`.

## VBA dependency declaration

VBA is now authorized. Every module must list the exact contract members it
uses (see `docs/vba-design.md`).

## Change control

After freezing, any change to a worksheet, Table, column, named range, status,
transaction type, formula contract, or Scan/Receiving interface location
requires:

1. documented problem;
2. proposed contract revision;
3. impact analysis;
4. owner approval;
5. contract version increase;
6. affected VBA and test updates;
7. full regression test.

## Deployment (NOT frozen)

`deployment_status: proposed` — the final storage/deployment option is an
unresolved owner decision (D-013 Proposed). Options A/B/C/D are documented in
`docs/architecture.md` §14. The frozen application architecture does not depend
on which option is finally chosen.
