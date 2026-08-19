# Production Deployment — Laboratory Inventory Excel v1.0.0

**Decision:** D-023 (Accepted, 2026-08-19) — supersedes D-013 (Proposed).
**Contract:** v1.0.0 frozen; deployment now Accepted (schema `deployment_status: accepted`, `deployment_option: option_a_network_share`).

## Model (Option A/B)

- **One dedicated Windows PC is the ONLY writer.** That PC is the designated
  laboratory inventory station.
- **All other PCs are read-only.** They use the macro-free report workbook —
  never the master `.xlsm`.
- **No simultaneous editing.** There is exactly one authoritative master and
  one writer; Excel's file lock + NTFS permissions make a second writer
  structurally impossible.

## Layout

| Role | Artifact | Location |
|---|---|---|
| Master (write) | `LabInventory_v1.0.0-production.xlsm` | `\\<LAB-SERVER>\Inventory\` (private SMB share) |
| Read-only report | `LabInventory_v1.0.0-readonly-report.xlsx` | `\\<LAB-SERVER>\Inventory\ReadOnly\` (or OneDrive shared folder for viewers) |
| Backups | `LabInventory_backup_<timestamp>.xlsm` | outside the live master share (e.g., `D:\InventoryBackups` on the server + archive) |

## Permissions

- Writer account (dedicated PC): **Modify** on the master share folder.
- All other accounts: **Read only** (no write, no delete, no rename).
- The report folder is **Read only** for viewers.

## Behavior details

- **File locking:** opening the master from the writer PC acquires Excel's
  exclusive lock; any other user who somehow obtains the master cannot open it
  for editing (read-only/denied). The report is a separate file, so viewers
  are never blocked by the master's lock.
- **Versioning:** timestamped backups (see `docs/backup-recovery.md`); Git
  versioning for source + release binaries.
- **Offline / network loss:** the master must be opened from the share. If the
  share is unreachable, do NOT copy the master to a local drive and work
  offline (would create a split-brain). The read-only report is a snapshot and
  opens without the share.
- **Failure/recovery:** restore = copy the newest verified backup to the
  master location; see `docs/backup-recovery.md` and
  `evidence/production/backup-restore-test.md`.
- **Read-only report refresh:** regenerate after each day's operations (or on
  demand) from the master's current state and replace the report file in the
  viewer location. Viewers always open the macro-free report.

## Why not Option C / D

- **Option C (OneDrive/SharePoint):** requires a Microsoft 365 business tenant
  with a SharePoint library and enforced permissions. The build machine runs
  Excel 2024 Retail with only a personal OneDrive — not available/appropriate
  without owner provisioning.
- **Option D (Google Drive sync):** sync introduces split-brain and is not a
  desktop-Excel/VBA runtime. Google Drive may be used only for backup/archive
  copies, never as the operational workbook.

## Owner notes

- Provision `\\<LAB-SERVER>\Inventory\` with the permission model above
  (replace `<LAB-SERVER>` with the actual server name).
- On first go-live, place the clean `LabInventory_v1.0.0-production.xlsm` on
  the share, import the initial inventory (see `docs/initial-data-import.md`),
  and record the go-live timestamp as the start of the electronic audit trail.
