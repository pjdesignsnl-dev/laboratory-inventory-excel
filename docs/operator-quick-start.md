# Operator Quick Start — Laboratory Inventory Excel v1.0.0

## Before you start

- Use the **dedicated writing PC** for all operations below.
- Open `LabInventory_v1.0.0-production.xlsm` from the network share
  (`\\<LAB-SERVER>\Inventory\`). The status bar must show
  **"Laboratory Inventory ready."** — this means the workbook's structure
  passed the frozen-contract check. If it shows a contract-violation message,
  stop and restore from backup (see `docs/backup-recovery.md`).
- **Do not** enable all macros globally and do not lower Trust Center
  security. The workbook must be trusted via a **Trusted Location** scoped to
  the production share folder (see `docs/scanner-configuration.md` /
  `docs/production-deployment.md`).

## Scan (daily use — the default entry point)

1. Go to the **Scan** sheet (the workbook opens to it).
2. Click/scan into the **Scan / type barcode** field (`Scan!D7`).
3. The status message below the field tells you:
   - `FOUND - scan details shown. Choose an action to commit.` — container
     exists; details and allowed actions are shown.
   - `UNKNOWN BARCODE - receive this container first.` — not in the system;
     receive it (below).
   - `Invalid barcode format (expected 7 digits).` — check the barcode.
   - `TakeOpen blocked: container is expired by date (D-018)...` — expired
     containers cannot be taken; record MarkExpired or Dispose.
4. After a successful scan, perform the action (Take/Open, Return, Transfer,
   Dispose, etc.). The status message confirms: `OK: <action> committed
   (T########).`
5. The input field clears and focus returns to it, ready for the next scan.

## Receiving a new container

1. On the **Receiving** sheet enter the Product ID, Batch/Lot, Expiry,
   Location, Quantity.
2. The Next Container ID and Next Barcode are shown.
3. Record the receive (the transaction is appended to the audit trail).
   - Initial stock is imported via the templates (see
     `docs/initial-data-import.md`) — receiving creates NEW containers after
     go-live.

## Reading stock (other PCs — no writing)

- Viewers open **`LabInventory_v1.0.0-readonly-report.xlsx`** (macro-free).
- It shows current stock, low/out-of-stock, expiry, reorder requirements,
  locations, and recent activity. It is refreshed by the writer PC after daily
  operations.
- Viewers never open the master `.xlsm`.

## Daily close

1. Verify the Dashboard has no error cells and totals look right.
2. On the writer PC: **File → Save** (or use the built-in backup; see
   `docs/backup-recovery.md`).
3. Regenerate and replace the read-only report for viewers.

## If something goes wrong

- Do not edit Transactions directly — corrections use the documented
  Adjustment workflow (compensating transaction).
- A failed operation rolls back automatically (no half-written data).
- For data loss/corruption, restore from the newest verified backup
  (`docs/backup-recovery.md`).
