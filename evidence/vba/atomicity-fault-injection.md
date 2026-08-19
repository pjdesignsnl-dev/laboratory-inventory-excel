# Atomicity under injected runtime failures — evidence

Date: 2026-08-18 (Microsoft Excel 16.0.20228.20190 x64, Windows)
Candidate: `workbook/LabInventory_v1.0-candidate.xlsm` (D-022 hash-free base,
24 VBA components incl. `modFaultInjection` test controller)
Full raw output: `evidence/vba/test-hooks-output.txt` (Test_FaultInjection)

## Purpose

Prove that every mutation boundary, when a deterministic runtime failure is
injected, leaves the workbook in the exact pre-operation state:

- no committed orphan transaction remains
- no partially mutated container remains
- original Container fields restored exactly
- transaction count / container count unchanged
- available-stock (Dashboard) unchanged
- formulas contain no errors
- protection/application state restored
- Application.EnableEvents restored to the pre-operation value
- calculation mode restored
- scanner interface remains usable
- a normal valid transaction succeeds immediately afterward

Rollback removes **only** rows belonging to the currently uncommitted failing
operation. Historical committed audit rows are never deleted.

## Fault-injection design (`modFaultInjection`)

One-shot armed faults raise error 2999 at a configured boundary on the Nth
matching call (`ArmFault(point, fireOnNthCall)`). Boundaries:

| # | Boundary | Raised by |
|---|---|---|
| 1 | before transaction append | `modTransactions.AppendTransaction` |
| 2 | immediately after transaction row allocation/append | `modTransactions.AppendTransaction` (self-cleaning: removes its own row, then raises) |
| 3 | before container mutation | `modContainers.ApplyStateChange` (scan-commit) / `modReceiving.ReceiveOne` before `AddContainer` (receive) |
| 4 | during/after partial container mutation | `modContainers.ApplyStateChange` (scan-commit) / `modContainers.AddContainer` (receive; self-cleaning) |
| 5 | after container mutation, before successful workflow completion | `modScanInterface.CommitAction` (scan-commit) / `modReceiving.ReceiveOne` (receive) |

Rollback correctness additions made for this acceptance:
- `AppendTransaction` removes its own just-created row if it raises after
  allocation (boundary 2 / post-condition), so no orphan can survive even when
  the caller has no tid handle.
- `AddContainer` removes its own just-created row if it raises mid-write
  (boundary 4 receive variant).
- `ReceiveOne` rollback now removes the Receive transaction it appended in
  addition to its container row.
- `ReceiveN` rollback now removes every batch member's Receive transaction
  (via `FindReceiveTransactionByContainer`) **and** every batch container row.
- `CommitAction` rollback already restored container fields + removed the
  transaction; boundary 5 now exercises that path.

## Results (all PASS in real Excel)

### Scan-commit path — C000001 (Available, LOC0004) → TakeOpen, fault at each boundary

| Boundary | op failed | txn count | container count | status restored | location restored | dashboard | formula errors | events restored | calc restored | scanner usable | input cleared | INVARIANTS |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | OK | 31→31 | 21→21 | OK | OK | 15→15 | 0 | OK | OK | OK | OK | **OK** |
| 2 | OK | 31→31 | 21→21 | OK | OK | 15→15 | 0 | OK | OK | OK | OK | **OK** |
| 3 | OK | 31→31 | 21→21 | OK | OK | 15→15 | 0 | OK | OK | OK | OK | **OK** |
| 4 | OK | 31→31 | 21→21 | OK | OK | 15→15 | 0 | OK | OK | OK | OK | **OK** |
| 5 | OK | 31→31 | 21→21 | OK | OK | 15→15 | 0 | OK | OK | OK | OK | **OK** |

Events assertion compares against the captured pre-operation value (the test
session runs with events disabled, so the correct "restored" value is False —
the invariant is equality with the pre-capture, proving no leak from the
operation itself).

### Receive path — ReceiveOne at each boundary

| Boundary | op failed | txn count | container count | dashboard | formula errors | events restored | calc restored | normal receive after |
|---|---|---|---|---|---|---|---|---|
| 1 | OK | 31→31 | 21→21 | 15→15 | 0 | OK | OK | OK (C000022) |
| 2 | OK | 31→31 | 21→21 | 15→15 | 0 | OK | OK | OK (C000022) |
| 3 | OK | 31→31 | 21→21 | 15→15 | 0 | OK | OK | OK (C000022) |
| 4 | OK | 31→31 | 21→21 | 15→15 | 0 | OK | OK | OK (C000022) |
| 5 | OK | 31→31 | 21→21 | 15→15 | 0 | OK | OK | OK (C000022) |

### ReceiveN batch rollback — fault on the 3rd of 5 members (2 fully created first)

| Check | Result |
|---|---|
| Batch failed (expected) | OK |
| No partial batch remains (container count) | 21→21 OK |
| No orphan Receive transactions (txn count) | 31→31 OK |
| Dashboard available unchanged | 15→15 OK |
| Formula errors | 0 OK |
| Events/calc restored | OK |
| Subsequent ReceiveN succeeds with unique IDs | OK — C000022,C000023,C000024 unique |

## Integrity notes

- After every injected failure, a **normal valid transaction succeeds
  immediately afterward** (scan re-scan → FOUND; ReceiveOne/ReceiveN succeed),
  proving the scanner interface and mutation engine remain fully usable.
- Historical committed audit rows are untouched: only the failing operation's
  own rows are removed (verified by unchanged transaction counts and by the
  append-only design).
