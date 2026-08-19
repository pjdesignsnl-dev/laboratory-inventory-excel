# Initial Data Import — Laboratory Inventory Excel v1.0.0

## Principle

- The production workbook ships **empty** (clean operational state, no
  synthetic fixtures). Initial stock is imported with these templates.
- **Never paste directly into protected/system columns.** Import through the
  validated templates below; the QA report records the checks performed.
- The **go-live timestamp** is the beginning of the electronic audit trail.
  Historical paper transactions are NOT migrated unless the lab provides them;
  no audit events are manufactured.
- If the actual laboratory inventory data is available, use the templates to
  capture it; do NOT invent data.

## Templates (CSV — in `templates/`)

1. `templates/import_products.csv` — Products catalogue.
2. `templates/import_suppliers.csv` — Suppliers.
3. `templates/import_locations.csv` — Storage locations.
4. `templates/import_containers.csv` — Initial container stock
   (ContainerID, Barcode, ProductID, Batch/Lot, ExpiryDate, RetestDate,
   DateReceived, StorageLocationID, Status, Notes).

## Required validation before import (QA report template: `templates/import_qa_report.md`)

For each file, run and record:

| Check | Rule |
|---|---|
| ProductID uniqueness | no duplicate P###### |
| ContainerID uniqueness | no duplicate C###### |
| Barcode uniqueness | no duplicate `\d{7}` text |
| FK references | Container.ProductID exists in Products; Container.StorageLocationID exists in Locations |
| Date types | ExpiryDate/RetestDate/DateReceived are real dates (or blank) |
| Mandatory fields | ProductID, ProductName, ContainerID, Barcode, Status present |
| Barcode format | exactly 7 digits, text |
| Expired stock | any Expired container is imported with Status=Expired (never counted as usable Available) |
| Status values | only the 6 frozen statuses |

## Import order

1. **Suppliers** (referenced by Products.SupplierID).
2. **Locations** (referenced by Containers.StorageLocationID).
3. **Products**.
4. **Containers** (initial stock).

## After import

- Open the master; Dashboard must show totals consistent with the imported
  data (usable available excludes expired-by-date).
- Take a backup (`LABINV_BACKUP_FOLDER` target) and record the go-live
  timestamp.
- Regenerate the read-only report.

## Migration of historical transactions

- Only if the lab supplies historical electronic records: import them as
  dated transactions using the Adjustment workflow with the original dates
  documented in Notes. Otherwise the audit trail starts at go-live.
