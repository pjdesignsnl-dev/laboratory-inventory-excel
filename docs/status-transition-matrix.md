# Container status and transaction transition matrix

**Status:** DRAFT — v0.1 (not frozen)
**Date:** 2026-08-17
**Contract reference:** `schema/workbook-contract.yaml`

This document is the authoritative state model for `LabInventory_v0.1.xlsx`. Every transition is classified:

- **ALLOW** — execute without additional warning;
- **CONFIRM** — show a clear warning and require confirmation;
- **BLOCK** — reject without mutation;
- **N/A** — transaction does not apply to this status.

## Status definitions

| Status | Meaning | In storage? | Qualifies as available stock? | Can be returned? | Notes |
|---|---|---|---|---|---|
| `Available` | Physically in a storage location, not opened, usable | Yes | **Yes** | N/A (already available) | The only stock-qualifying status |
| `InUse` | Taken/opened out of storage, in active use | No | No | Yes (via `Return`) | Remains traceable; `OpenedDate` set |
| `Reserved` | Held for a specific purpose, still physical | Yes | No | — | Reserved does not consume stock |
| `Expired` | Past expiry, unusable | Yes/No | No | No | Requires `Dispose` to leave inventory |
| `Damaged` | Physically unusable | Yes/No | No | No | Requires `Dispose` |
| `Disposed` | Removed from inventory (used up, expired, damaged, etc.) | No | No | No | Terminal; `DisposalDate`/`DisposalReason` set |
| `Missing` | Cannot be located; suspected lost | Unknown | No | No | Requires `Adjustment`/`Dispose` resolution |

**Flag semantics:** `OpenedDate` on a Container is independent of status. An `Available` container with `OpenedDate` set is a returned, previously opened container (default assumption 12; decision D-005). This is why `Opened/In Use` is **not** a status.

## Transaction definitions

| Transaction | Effect on Container | Required fields | Warning/confirm | Stock effect |
|---|---|---|---|---|
| `Receive` | Creates container; Status → `Available` (default `DefaultStatusNewContainers`); sets DateReceived; sets location | Barcode, ProductID, lot, location; expiry optional | CONFIRM if barcode already exists (blocked); info otherwise | +1 |
| `TakeOpen` | Status → `InUse`; sets `OpenedDate` (if unset); clears location | None beyond scan | CONFIRM if already opened or if expiry soon | −1 (leaves Available) |
| `Return` | Status → `Available`; sets location; keeps `OpenedDate` | Location | CONFIRM if expired/damaged/missing (should be blocked in VBA) | +1 |
| `Transfer` | Location changes; status unchanged | New location | ALLOW | 0 |
| `Dispose` | Status → `Disposed`; sets DisposalDate + DisposalReason; clears location | DisposalReason | **CONFIRM** (destructive, irreversible) | −1 if was Available |
| `MarkExpired` | Status → `Expired` | None | CONFIRM (auto-set on scan of expired in VBA) | −1 if was Available |
| `MarkDamaged` | Status → `Damaged` | Reason | CONFIRM | −1 if was Available |
| `MarkMissing` | Status → `Missing` | Reason | CONFIRM | −1 if was Available |
| `Adjustment` | Status/location reconciliation; always paired with reason | Reason, previous/new state snapshots | **CONFIRM** | ±1 per direction |

## Matrix

| Current status | Receive | Take/Open | Return | Transfer | Dispose | Mark Expired | Mark Damaged | Mark Missing | Adjustment/Reversal |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| (none — new) | **ALLOW** | N/A | N/A | N/A | N/A | N/A | N/A | N/A | N/A |
| `Available` | BLOCK | **ALLOW** | N/A | **ALLOW** | **CONFIRM** | **CONFIRM** | **CONFIRM** | **CONFIRM** | CONFIRM |
| `InUse` | BLOCK | BLOCK (already opened) | **ALLOW** | BLOCK (not in storage) | CONFIRM | CONFIRM | CONFIRM | CONFIRM | CONFIRM |
| `Reserved` | BLOCK | CONFIRM (release + take) | ALLOW | ALLOW | CONFIRM | CONFIRM | CONFIRM | CONFIRM | CONFIRM |
| `Expired` | BLOCK | BLOCK (invariant 9) | BLOCK | ALLOW (relocation) | **ALLOW** | N/A (already) | BLOCK (expired) | BLOCK | CONFIRM |
| `Damaged` | BLOCK | BLOCK | BLOCK | ALLOW | **ALLOW** | BLOCK | N/A (already) | BLOCK | CONFIRM |
| `Disposed` | BLOCK | BLOCK | BLOCK | BLOCK | BLOCK (already) | BLOCK | BLOCK | BLOCK | CONFIRM (reopen/restore) |
| `Missing` | BLOCK | BLOCK | BLOCK | BLOCK | CONFIRM (resolve as disposed) | BLOCK | BLOCK | N/A (already) | CONFIRM (found → Available) |

**Invariant enforcement:** expired/damaged/disposed/missing containers cannot be taken through a normal available-stock transaction (invariant 9). The `Adjustment` row is the only path to reverse a terminal state, and it always requires explicit confirmation and a reason.

## Side effects

### Receive
- Resulting status: `Available` (or `DefaultStatusNewContainers`).
- Location: set from entry; defaults to `DefaultLocationID`.
- Dates: `DateReceived` = today; `OpenedDate`/`DisposalDate` blank.
- Required: barcode unique, product, lot, location.
- Snapshot: full product/batch/status/location snapshot; `PreviousStatus=(none)`, `PreviousLocation=(none)`.
- Stock effect: +1 available.
- Reversal: `Adjustment` (mark disposed) or `Dispose`.

### TakeOpen
- Resulting status: `InUse`.
- Location: cleared (not in storage).
- Dates: `OpenedDate` set if blank.
- Warning: CONFIRM if `OpenedDate` already set (re-take) or expiry soon.
- Snapshot: previous `Available`, new `InUse`, location cleared.
- Stock effect: −1 available.
- Reversal: `Return`.

### Return
- Resulting status: `Available`.
- Location: set to entry location.
- Dates: `OpenedDate` preserved.
- Warning: BLOCK if terminal status; CONFIRM if expiry soon.
- Stock effect: +1 available.
- Reversal: `TakeOpen` (re-take).

### Transfer
- Resulting status: unchanged.
- Location: new location.
- Warning: ALLOW.
- Stock effect: 0.
- Reversal: `Transfer` back.

### Dispose
- Resulting status: `Disposed`.
- Location: cleared.
- Dates: `DisposalDate` = today; `DisposalReason` set.
- Warning: CONFIRM (irreversible in normal operation).
- Stock effect: −1 if was Available.
- Reversal: `Adjustment` (restore) — requires reason.

### MarkExpired / MarkDamaged / MarkMissing
- Resulting status: `Expired` / `Damaged` / `Missing`.
- Location: kept for Expired/Damaged; cleared for Missing.
- Warning: CONFIRM.
- Stock effect: −1 if was Available.
- Reversal: `Adjustment` only.

### Adjustment
- Resulting status: any valid target.
- Location/dates: as set in entry.
- Warning: CONFIRM, requires `Reason`.
- Stock effect: ±1 per direction.
- Purpose: corrections use compensating events (invariant 7); never overwrite history.
