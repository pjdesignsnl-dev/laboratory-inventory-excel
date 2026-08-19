# Backup & Recovery — Laboratory Inventory Excel v1.0.0

## Backup destination

- **Default (in-workbook):** `modBackup.BackupFolder()` uses the
  `LABINV_BACKUP_FOLDER` environment variable if set, otherwise the workbook's
  own folder.
- **Production:** set `LABINV_BACKUP_FOLDER` on the writer PC to a folder
  **outside the live master share**, e.g. `D:\InventoryBackups` (local disk on
  the server), mirrored to an archive location (second disk or Google Drive
  archive copy — Google Drive is acceptable for archive/backup only, never as
  the operational workbook).

## Behavior (verified)

- `modBackup.CreateBackup` writes `LabInventory_backup_<yyyymmdd_hhnnss>.xlsm`
  (never overwrites — if a name collides, an `_N` suffix is appended).
- A required backup that fails **stops the operation** (no silent continue).
- Recovery is operator-driven: the newest verified backup is copied to the
  master location.

## Retention recommendation

- **Daily operational backups:** keep 14 (14 days).
- **Weekly retained copies:** keep 8 (2 months).
- **Monthly archive:** keep 12 (12 months).
- Keep at least one backup on a different physical disk than the master.

## Restore drill (performed 2026-08-19)

See `evidence/production/backup-restore-test.md` for the executed drill:
backup → SHA/date recorded → restore to a test location → open → contract
validates → formulas calculate → macros work → scan works → master unchanged.

## Manual restore steps (operator)

1. On the writer PC, open the newest backup
   (`LabInventory_backup_<latest>.xlsm`) and confirm:
   - status bar shows "Laboratory Inventory ready." (contract OK),
   - Dashboard recalculates with no error cells,
   - a test scan of a known barcode shows FOUND.
2. Close it. Copy it over the master file name on the share.
3. Open the restored master from the share and repeat step 1.
4. Record the restore in the backup log.
