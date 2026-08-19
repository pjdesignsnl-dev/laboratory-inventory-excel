# Initial Data Import — QA Report

**Import run by:** __________ **Date:** __________
**Go-live timestamp (start of electronic audit trail):** __________

## Files imported

| File | Rows | ProductID unique | ContainerID unique | Barcode unique | FK ok | Dates ok | Mandatory ok | Barcode fmt ok | Status ok | Notes |
|---|---|---|---|---|---|---|---|---|---|---|
| import_suppliers.csv | | n/a | n/a | n/a | n/a | n/a | | n/a | n/a | |
| import_locations.csv | | n/a | n/a | n/a | n/a | n/a | | n/a | n/a | |
| import_products.csv | | | n/a | n/a | Suppliers | n/a | | n/a | n/a | |
| import_containers.csv | | n/a | | | Products+Locations | | | | | |

## Expired-stock check

- Any container with ExpiryDate < today must be imported with Status=Expired
  (it must NOT count toward usable Available). Count imported as Expired: ____
- Dashboard "Total available containers (usable)" after import: ____

## Baseline audit

- Production master SHA-256 after import: __________
- Backup created: `LabInventory_backup_<timestamp>.xlsm` at: __________
- Read-only report regenerated and placed at: __________

## Sign-off

Writer operator: __________  Reviewer: __________

(Attach the import files and this report to the go-live record.)
