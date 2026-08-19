## Deployment-location smoke — 2026-08-19 23:11:29

**Simulated deployment location:** C:\Users\Q\Documents\laboratory-inventory-excel\evidence\production\deployment-location (stands in for the D-023 network share \\\\<LAB-SERVER>\\Inventory\\)
**Deployed master SHA-256:** A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1
**Committed clean production master SHA-256 (must stay unchanged):** A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1

### 1. Open from deployment location
- Workbook_Open contract validation (status bar): [Laboratory Inventory ready.]
- Contract OK: True
- Driver imported (throwaway, not part of any committed binary)
- driver-start
- import-minimal-reference:OK
- receive-one=[C000001]
- receive-one-ok=True
- takeopen=[OK: TakeOpen committed (T00000002).]
- return=[OK: Return committed (T00000003).]
- txn-count=3
- status=Available
- driver-done
### 2. Operator flow results
- Transaction rows appended: 3 (expect 3: Receive, TakeOpen, Return)
- Dashboard error cells: 0 (0 expected)
- Backup created (timestamped, no overwrite): C:\Users\Q\Documents\laboratory-inventory-excel\evidence\production\deployment-location\backups\LabInventory_backup_20260819_231139.xlsm
- Reopen: containers=1 (expect 1) — persistence OK: True
- Save/close/reopen persistence: PASS
### 3. Second writer not permitted
- Second instance opened the deployed master: True
- Second writer permitted: False

### 4. Committed clean production master unchanged
- Pre-smoke SHA: A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1
- Post-smoke SHA: A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1
- Master unchanged: True

**AccessVBOM restored to:** 0
