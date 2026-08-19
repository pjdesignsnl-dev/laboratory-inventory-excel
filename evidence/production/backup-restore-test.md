## Backup/restore drill — 2026-08-19 22:39:59

**Production master:** C:\Users\Q\Documents\laboratory-inventory-excel\workbook\LabInventory_v1.0.0-production.xlsm
**Pre-drill master SHA-256:** A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1

### 1. Backup created
- Path: C:\Users\Q\Documents\laboratory-inventory-excel\evidence\production\restore-test\LabInventory_backup_20260819_223959.xlsm
- SHA-256: A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1
- Timestamp: 20260819_223959 (no overwrite; unique timestamped name)

### 2. Restore to test location
- Restored to: C:\Users\Q\Documents\laboratory-inventory-excel\evidence\production\restore-test\restored-master.xlsm
- Restored SHA-256: A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1 (must equal backup SHA)
- Restore SHA matches backup: True

### 3. Restored workbook opened in real Excel
- Status bar: [Laboratory Inventory ready.]
- Contract validates (ready): True
- Dashboard error cells: 0 (0 expected)
- Scan macro responds: [UNKNOWN BARCODE - receive this container first.]
- Scan works: True

### 4. Production master unchanged
- Post-drill master SHA-256: A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1
- Master unchanged: True

**Result:** PASS
