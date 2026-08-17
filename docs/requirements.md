# Authoritative product requirements

## 1. Outcome

Design and build a professional Excel-based laboratory inventory system for consumables and chemicals. It replaces a monthly paper-entry process with a barcode-driven workflow that is reliable, simple, traceable, maintainable, and scalable.

Treat it as a small inventory-management application inside Excel, not a generic template.

Priority order:

1. Data integrity
2. Individual physical-container traceability
3. Reliable barcode scanning
4. Simple user experience
5. Permanent audit history
6. Error prevention
7. Chemical expiry tracking
8. Maintainability
9. Scalability
10. Professional appearance

## 2. Required incremental development sequence

Do not design the entire workbook and VBA simultaneously. Finalize and validate workbook architecture, data model, Tables, relationships, controlled values, formulas, interfaces, and workflows before writing VBA.

Required stages:

1. Requirements and assumptions
2. Architecture
3. Architecture review
4. Workbook construction
5. Non-VBA architecture testing
6. VBA design against the exact frozen workbook
7. Incremental VBA implementation and module testing
8. Integration testing
9. Final deployment documentation

If VBA development reveals an architectural problem, stop, explain it, propose a contract change, and wait for approval. Never silently change names or structures.

## 3. Environment

- Windows desktop
- Microsoft 365 Excel and/or Excel 2021/2024
- VBA/macros permitted
- One dedicated Windows PC in or near the laboratory store
- Standard USB keyboard-wedge barcode scanner
- Dedicated PC is the only inventory writer
- Other PCs require read-only access
- Multiple simultaneous editors are unnecessary

## 4. Core scan workflow

The operational flow must be fast enough for repeated use without navigating Excel manually:

1. Scanner ready
2. Scan barcode
3. Excel receives barcode
4. Identify physical container
5. Display product/container details and current status
6. Select transaction type
7. Validate and request required confirmation
8. Permanently append transaction
9. Update current container status/location
10. Recalculate inventory/dashboard
11. Display clear success, warning, or error
12. Clear scan field
13. Return focus to barcode input
14. Ready for next scan

Minimize mouse use. Recommend scanner configuration, including an Enter suffix.

## 5. Inventory granularity

Every physical container/package eventually has its own unique barcode. Identical products still have separate Container records and barcodes.

Examples:

- Bottle A -> one Container and one barcode
- Bottle B -> another Container and barcode
- One box of 96 pipette tips -> one Container
- One box of 100 tubes -> one Container
- One package of filters -> one Container

Do not track remaining chemical volume, mass, or individual consumable pieces in v1. If 100 mL is used from a 500 mL bottle, it remains one physical container. Record that it was opened/used, but do not calculate remaining volume.

## 6. Relational-style concepts

At minimum, separate:

### Product

The general catalogue item, for example Ethanol 96%, laboratory grade, Manufacturer X, catalogue ABC123.

### Container

One physical instance, for example:

- Container ID C000123
- Barcode 1000001
- Product reference
- Lot LOT12345
- Expiry 2028-04-15
- Received 2026-08-10
- Location Chemical Cabinet 2
- Status Available

### Transaction

One immutable historical event containing the prior and resulting state. A barcode search must reconstruct the complete container history.

The architecture should also support Suppliers, Locations, Settings/controlled lists, scanning, receiving, dashboard/reporting, and optional batch-level separation when justified.

## 7. Product information

Design exact mandatory and optional fields. Consider:

- ProductID
- ProductName
- ProductType
- Category
- Manufacturer
- ManufacturerCatalogueNumber
- CASNumber
- Concentration
- Grade/Purity
- StandardContainerDescription
- SupplierID
- StorageRequirements
- HazardClassification
- SDSReference
- MinimumContainerStock
- TargetContainerStock
- ReorderQuantity
- Active
- Notes

Do not force chemical-specific values onto consumables. Distinguish product-level, batch-level, and container-level data.

## 8. Container information

Design exact fields. Consider:

- ContainerID
- Barcode
- ProductID
- Batch/LotNumber
- ExpiryDate
- RetestDate
- DateReceived
- StorageLocationID
- Status
- OpenedDate
- DisposalDate
- DisposalReason
- Notes

A manually maintained CurrentQuantity is not required because stock is container-based and remaining contents are not tracked.

## 9. Permanent transaction history

Use an append-only Excel Table. At minimum consider:

- TransactionID
- Timestamp
- User/Operator (reserved for future use)
- Barcode
- ContainerID
- ProductID
- ProductName snapshot
- TransactionType
- PreviousStatus
- NewStatus
- PreviousLocation
- NewLocation
- Batch/Lot snapshot
- Reason
- Reference
- Notes

Add other fields when they improve integrity or forensic reconstruction. Normal operation must never delete or overwrite historical transactions. Corrections should use compensating/reversal events.

## 10. Transaction types and effects

Support and precisely define, where appropriate:

- Receive
- Take/Open
- Return
- Transfer
- Dispose
- Expired
- Damaged
- Lost/Missing
- Stock adjustment/correction

Prevent invalid state transitions. Examples:

- Receive creates/activates a physical Container and logs it.
- Take/Open removes a currently available container from storage availability and records opening/use.
- Return restores a valid container to a storage location.
- Transfer changes a storage location while preserving history.
- Dispose preserves the Container and history but removes it from active inventory.
- Expired, Damaged, and Missing make the Container unavailable for normal use.

Define warning-versus-blocking behavior explicitly.

## 11. Status model

Develop the smallest practical model. Candidate concepts include:

- Available
- Opened/In Use
- Taken
- Reserved
- Expired
- Damaged
- Disposed
- Lost/Missing

Do not create redundant states. Clearly distinguish:

- physically in storage;
- previously opened;
- currently taken/in use;
- empty;
- disposed.

Because contents are not measured, the system cannot infer empty; the user must record an appropriate disposal/empty event.

Consider representing `OpenedDate` independently from current availability so that an opened bottle can be returned and available while remaining historically identified as opened.

## 12. Stock calculation

Stock is based on counts of physical Containers that qualify as available.

Examples:

- Ethanol: 3 available containers
- Pipette tips: 12 available boxes
- Laboratory tubes: 7 available boxes

Taking one available bottle changes 3 -> 2. Returning it changes 2 -> 3. Disposing of an available bottle changes 3 -> 2.

Choose between dynamic calculation from Containers, event reconstruction from Transactions, stored balances, or a controlled combination. Prioritize integrity. A separately editable stock number is undesirable.

## 13. Chemicals

Support:

- CAS number
- concentration
- grade/purity
- manufacturer and catalogue number
- batch/lot
- expiry date
- retest date where applicable
- hazard classification
- storage requirements
- SDS reference

Provide configurable expiry warning bands, initially 30/60/90 days. Distinguish expired, expiring soon, and valid containers.

This workbook is not a replacement for formal chemical safety, SDS, regulatory, validated LIMS, GMP/GLP, or hazardous-material systems.

## 14. Consumables

Support complete packages such as pipette tips, tubes, filters, plates, syringes, gloves, and other disposable laboratory materials. Taking one complete package decreases available stock by one. Do not track individual contents.

## 15. Reordering

Per Product support:

- Minimum container stock
- Target container stock
- Reorder quantity

Identify low stock, out of stock, and reorder required based on available physical Containers.

## 16. Validation and messaging

Detect at minimum:

- unknown barcode;
- duplicate barcode;
- barcode assigned to multiple Containers;
- invalid barcode;
- accidental duplicate scan;
- disposed container;
- expired container;
- missing container;
- invalid transaction for current state;
- missing required data;
- invalid product/location/reference.

Classify validation results as:

- blocking error;
- warning requiring confirmation;
- informational message.

Invalid or failed transactions must not alter inventory or create partial history.

## 17. Receiving

Create an efficient receiving workflow:

1. Select Product
2. Generate immutable ContainerID
3. Scan or enter unique barcode
4. Enter lot/batch
5. Enter expiry/retest dates when applicable
6. Select storage location
7. Set valid initial state
8. Append Receive transaction

Make batches of identical bottles efficient while still assigning one ContainerID/barcode to every physical unit.

Evaluate internal generated labels versus pre-existing supplier labels. Manufacturer product barcodes cannot uniquely identify multiple identical physical Containers, so a unique internal barcode is generally required. Recommend label generation/printing method and Code 128 or another suitable symbology.

## 18. Dashboard

Prioritize operational usefulness over decoration. Include:

- total active products;
- total available containers;
- low-stock and out-of-stock products;
- expired containers;
- containers expiring soon;
- recently received containers;
- recently taken/used containers;
- frequently used products;
- inventory by category;
- inventory by location;
- reorder requirements.

## 19. Workbook architecture

A suggested starting set is:

- Dashboard
- Scan
- Receiving
- Products
- Containers
- Transactions
- Suppliers
- Locations
- Settings

The architecture may change if a demonstrably better design is documented. For every worksheet specify purpose, editor/viewer, Tables, relationships, protected areas, and intended workflow.

## 20. Read-only access and deployment architecture

Compare:

A. One master `.xlsm` on a network location
B. Master `.xlsm` plus a separate read-only reporting workbook
C. Microsoft 365/SharePoint/OneDrive architecture
D. Google Drive synchronized Office-file storage where relevant

Discuss desktop VBA compatibility, file locking, network/sync failures, corruption risk, backups, versioning, read-only access, and accidental editing. Recommend one primary architecture for one writer and multiple readers.

Do not convert the operational `.xlsm` to native Google Sheets.

## 21. Formula and workbook requirements

Use Excel Tables and structured references. Provide exact compatible formulas for:

- available container count;
- low-stock and out-of-stock detection;
- reorder requirements;
- expiry warning bands;
- duplicate barcode detection;
- product/container lookup;
- dashboard statistics.

Prefer Excel 2021/2024 compatibility and identify Microsoft 365-only alternatives. Configure controlled lists, validation, named ranges when useful, conditional formatting, appropriate sheet/workbook protection, and a polished Scan interface.

## 22. VBA design and implementation rules

VBA is the final implementation layer. Before VBA, finalize and test:

- worksheet names;
- Table names;
- column names;
- IDs;
- statuses;
- transaction types;
- Scan-screen ranges;
- named ranges;
- formulas;
- validation.

For every VBA feature explicitly identify all workbook dependencies. Use exact frozen names consistently.

Implement incrementally in logical modules such as:

1. Barcode input and lookup
2. Container/transaction validation
3. Atomic transaction logging
4. Container state/location updates
5. Receiving
6. Scan interface reset/focus/messages
7. Error handling and diagnostics
8. Backup/recovery

For every module provide complete usable code, exact destination, dependencies, trigger, workbook interactions, test steps, expected results, and failure cases. Avoid unnecessary `Select`/`Activate`, fragile references, and hard-coded data rows. Use constants/enumerations where appropriate.

Atomicity is mandatory: a failure must not leave a half-written Transaction or partially updated Container.

## 23. Testing

Testing is incremental:

- test architecture;
- test formulas and validation;
- test each VBA module independently;
- run full integration testing;
- run physical scanner acceptance testing.

Normal scenarios include:

- receive new container;
- scan available container;
- take/open;
- return;
- transfer;
- dispose.

Error scenarios include:

- unknown/duplicate barcode;
- accidental duplicate scan;
- disposed or expired container;
- invalid transition;
- missing required expiry;
- invalid product/location;
- network/file problem;
- VBA error;
- user cancellation.

Data-integrity acceptance must prove:

- Transactions are never overwritten;
- complete Container history remains reconstructable;
- stock counts remain correct;
- invalid transactions make no changes;
- failures do not create half-transactions;
- duplicate barcodes cannot exist.

The end-to-end workflow to prove is:

`Scan -> identify -> validate -> select transaction -> confirm -> append transaction -> update container -> update stock -> display result -> reset -> focus for next scan`.

## 24. Protection, backups, and limitations

Use locked formula cells, protected worksheets, protected workbook structure, controlled dropdowns, restricted manual editing, configuration hiding where appropriate, append-only workflow enforcement, backups, recovery, and release versioning.

Explain that Excel protection is primarily an accidental-change deterrent, not strong security. Provide backup frequency, retention, recovery drills, and corruption handling.

## 25. Worked examples

Use synthetic examples for:

1. Box of pipette tips
2. Box of laboratory tubes
3. Bottle of ethanol
4. Another laboratory solvent
5. Reagent with batch and expiry

Demonstrate Product creation, Container registration, barcode assignment, receiving, scanning, taking/opening, returning, expiry warning, disposal, transaction history, and stock calculation.

## 26. Final delivery documentation

The completed project must include:

1. Requirements and assumptions
2. Final architecture
3. Final workbook/worksheet structure
4. Final Tables and relationships
5. Final user interface
6. Formulas and validation
7. VBA modules and exported source
8. Barcode scanner configuration
9. Testing and acceptance evidence
10. Deployment
11. Backup and maintenance
12. Future expansion

The owner must receive enough information and files to build, deploy, operate, test, recover, and maintain the workbook—not merely a conceptual description.
