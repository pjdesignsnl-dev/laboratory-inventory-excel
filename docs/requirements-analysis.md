# Requirements and assumptions analysis

**Status:** COMPLETE — v0.1 architecture basis
**Date:** 2026-08-17
**Task:** Stage 1 of `docs/INITIAL_TASK.md`
**Inputs:** `AGENTS.md`, `docs/requirements.md`, `docs/default-assumptions.md`

## 1. Environment inspection (actual, recorded)

The build machine is a Windows desktop that hosts this repository. Recorded facts:

| Item | Value |
|---|---|
| Operating system | Windows 10 IoT Enterprise LTSC 2021 (21H2), build 19044, AMD64 |
| Microsoft Office / Excel | **Not installed.** No `EXCEL.EXE`, Click-to-Run registry keys, Appx package, or WPS/LibreOffice installation was found. Excel edition/build is therefore **not discoverable** without changing system configuration. |
| Python | 3.11.0 (64-bit) at `C:\Users\Q\AppData\Local\Programs\Python\Python311` |
| openpyxl | 3.1.5 (installed workspace-locally into `.tools/pylib`, not system-wide) |
| formulas library | 1.3.4 (workspace-local) — used for non-VBA formula evaluation |
| Pillow | 12.3.0 (workspace-local) |
| 7-Zip | present at `C:\Program Files\7-Zip\7z.exe` |
| LibreOffice (portable, headless) | download attempted for verification/rendering; network-dependent |

**Consequence (recorded, not hidden):** desktop Excel verification **could not be executed on this machine**. Per `AGENTS.md` ("Do not claim desktop Excel/VBA/scanner tests passed unless they actually ran on Windows desktop Excel"), all tests in this phase are **structural inspection plus independent formula evaluation** using openpyxl/formulas, and (if the portable renderer is obtainable) a LibreOffice headless open/recalc/export smoke test. Desktop Excel verification remains a required owner-side step before the contract is frozen for VBA. This is logged as decision D-004.

## 2. Assumptions applied (from `docs/default-assumptions.md`)

All 18 defaults are applied; none were contradicted by `docs/requirements.md`. Status of each:

| # | Assumption | Applied | Notes |
|---|---|---|---|
| 1 | English UI | Yes | All labels, lists, and docs in English |
| 2 | PascalCase stable column names; `tbl`/`rng` prefixes | Yes | Exact names fixed in `docs/architecture.md` and `schema/workbook-contract.yaml` |
| 3 | `.xlsx` macro-free first build | Yes | `workbook/LabInventory_v0.1.xlsx` |
| 4 | One authoritative editor | Yes | Documented in deployment section |
| 5 | Read-only reporting copy later | Yes | Documented as future artifact |
| 6 | Internal numeric barcodes, Code 128, Enter suffix | Yes | Scanner configuration documented |
| 7 | Barcode stored as text | Yes | Cell number format `@`; leading zeroes preserved |
| 8 | ContainerID distinct from barcode | Yes | `C######` vs 7-digit barcode |
| 9 | TransactionID immutable; append-only | Yes | Column + convention + protection design |
| 10 | Stock = dynamic count of available Containers | Yes | No editable stock balance |
| 11 | One container = one stock unit | Yes | |
| 12 | Opened via OpenedDate + history; may return to Available | Yes | `Opened` is a flag, **not** a status |
| 13 | No login; Operator field reserved | Yes | `Operator` column on Transactions |
| 14 | 30/60/90-day expiry thresholds | Yes | Stored in `tblSettings` |
| 15 | Excel 2021/2024 + Microsoft 365 | Yes | All formulas 2021-compatible; no `FILTER`/`XLOOKUP`/`SORT` |
| 16 | Synthetic sample data | Yes | 5+ products, 20+ containers, 30+ transactions |
| 17 | Compliance boundary | Yes | Stated on Dashboard and in docs |
| 18 | Stop before VBA for review | Yes | Stop condition honored |

## 3. Requirement interpretation and key design resolutions

### 3.1 Granularity and quantity model (Req §5, §12)
One physical container/package = one inventory unit. Remaining mL/g/pieces are not tracked (invariant 6). Empty cannot be inferred; it requires a recorded `Dispose`/`Adjustment` event (Req §11). Stock is derived exclusively by counting Containers in qualifying available states (invariant 5).

### 3.2 Status model (Req §11)
Statuses chosen: `Available`, `InUse`, `Expired`, `Damaged`, `Disposed`, `Missing` (6 values, smallest practical v1 set — decision D-016). `Reserved` was removed because no complete reservation workflow exists in v1 and a state only practically enterable through `Adjustment` must not be retained. `OpenedDate` is an independent flag: a returned opened container can be `Available` again while retaining opened history (default assumption 12, decision D-005). This cleanly separates "physically in storage" from "previously opened" without extra states. Available-stock semantics (D-018): `Status="Available" AND (ExpiryDate blank OR ExpiryDate >= TODAY())` — an expired-by-date container is excluded from usable stock without silently mutating its stored Status.

### 3.3 Transaction model (Req §9, §10)
Types: `Receive`, `TakeOpen`, `Return`, `Transfer`, `Dispose`, `MarkExpired`, `MarkDamaged`, `MarkMissing`, `Adjustment` (9 values). Every transaction snapshots barcode, ContainerID, ProductID, ProductName, batch/lot, previous/new status, previous/new location, timestamp, and optional reason/reference/notes — satisfying "a barcode search must reconstruct the complete container history" (Req §6). Corrections are compensating `Adjustment` transactions (append-only; invariant 7).

### 3.4 Product vs batch vs container placement (Req §7, §8, §13)
- **Product level** (`tblProducts`): catalogue identity, manufacturer/catalogue/CAS/concentration/grade, hazard/storage/SDS, stock thresholds, reorder quantity.
- **Container level** (`tblContainers`): barcode, lot/batch, expiry/retest, received date, location, status, opened/disposal dates, notes.
- **Batch/lot level**: v1 has no separate batch table; lot is a Container attribute, because a container is a lot-level physical unit. Optional batch-level separation is deferred (documented risk).

### 3.5 Barcode design (Req §17, §20; default 6–8)
Unique internal 7-digit numeric barcode stored as **text**; manufacturer codes cannot uniquely identify identical containers. Code 128 recommended for labels; scanner configured as keyboard wedge with **Enter suffix**. Duplicate detection is formula-driven and also enforced by VBA later (invariant 10).

### 3.6 Stock / reorder / expiry formulas (Req §12, §13, §15)
- Available stock per product: `COUNTIFS(tblContainers[ProductID], <pid>, tblContainers[Status], "Available")`.
- Classification: `OutOfStock` if 0, `Reorder` if below `MinimumContainerStock`, `Low` if below `TargetContainerStock`, else `OK`.
- Expiry class: `Expired` (date < TODAY), `Expiring≤30/60/90`, `Valid`, `None` (no date). Bands from `tblSettings`.
- All formulas Excel-2021 compatible (decision D-012).

### 3.7 Scan workflow without VBA (Req §4, §21)
The Scan sheet is fully formula-driven for **lookup and validation display**: scanning (or typing) a barcode shows container, product, status, location, expiry class, duplicate flag, and the allowed next actions for the current status. Committing a transaction requires either manual table entry (interim v0.1 mode, heavily guarded by validation) or, after contract freeze, the VBA layer. The exact named ranges are fixed now so VBA does not need renames (Req §22).

### 3.8 Receiving without VBA (Req §17)
The Receiving sheet is an **interface foundation**: product picker, auto-generated next IDs, barcode/lot/expiry/location entry block, batch-size field, and step-by-step instructions. Non-VBA mode requires manual append of Container + `Receive` transaction rows; VBA later performs the atomic append. This is explicitly an interim manual workflow, not hidden automation.

### 3.9 Deployment (Req §20)
Recommendation: **Option A — one master `.xlsm` on a network share with a single writer PC** (default), with a published read-only copy (Option B) for other PCs once the workbook is stable. Full trade-off analysis in `docs/architecture.md` §14. Decision D-013 is **Proposed** (not Accepted): it is listed in the owner-decision report as an unresolved decision requiring owner confirmation before the contract freeze.

## 4. Ambiguities resolved (with default-based resolution)

| Ambiguity | Resolution |
|---|---|
| "Opened/In Use" as status or flag? | Flag (`OpenedDate`) + `InUse` status; return-to-Available allowed |
| Are `Expired`/`Damaged`/`Missing` transaction types or statuses? | Both: statuses, plus explicit marking transaction types (`MarkExpired`, `MarkDamaged`, `MarkMissing`) |
| Is empty content a state? | No — cannot be inferred; `Dispose`/`Adjustment` event required |
| Label printing in scope for v0.1? | Out of scope; barcode rules + symbology recommendation only |
| Operator capture? | Column reserved; Windows-username capture deferred to VBA phase |
| Is a batch/lot table needed? | Not in v1; lot is a Container attribute |

## 5. Risks and mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| No Excel on build machine → no desktop verification in this phase | Medium | Structural inspection + independent formula evaluation + (if obtainable) LibreOffice headless smoke test; owner desktop verification step defined; honestly labeled evidence |
| Manual-edit interim mode can violate atomicity | High | Validation, conditional formatting, protection on control sheets, explicit user instructions; VBA-phase enforcement is contract-mandated |
| Protection is deterrent, not security (Req §24) | Medium | Documented; backups + append-only convention |
| Formula columns duplicate lookups (integrity risk) | Medium | Lookups are formula-driven from source Tables, never hand-edited |
| Excel version differences | Low | 2021-compatible formulas only; 365-only alternatives documented as notes |
| Large tables → performance | Low | v0.1 fixture scale is small; COUNTIFS/COUNTIF-based |

## 6. Open questions for owner (non-blocking, needed before contract freeze)

1. **Deployment primary**: confirm Option A (single master `.xlsm` on network share) versus SharePoint/OneDrive (Option C), given the "one writer" constraint. Default assumes A.
2. **Label printing**: is internal label generation/printing (Code 128) in scope for v1, or only barcode assignment rules? Default: rules + symbology recommendation only.
3. **Operator capture**: capture Windows username at transaction time in VBA phase? Default: yes, in the reserved `Operator` column.
4. **Status set**: approve the 6-value status model (`Available`, `InUse`, `Expired`, `Damaged`, `Disposed`, `Missing`; `Opened` is a flag; `Reserved` removed) before freezing.
5. **Receiving batch efficiency**: is a "receive N identical containers" multi-row helper desired in v1 (VBA phase), or is one-by-one acceptable? Default: batch helper planned.
6. **Retest dates**: keep `RetestDate` as an optional advisory column only (no retest workflow in v1)? Default: yes.

None of these block Stages 1–5; they gate the contract freeze and the VBA phase.
