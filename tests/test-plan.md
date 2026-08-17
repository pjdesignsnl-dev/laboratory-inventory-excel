# Incremental test plan

## Test evidence rule

A test is not `PASS` unless the actual artifact was exercised in the stated environment and evidence was recorded. Do not infer VBA or scanner success from code review alone.

## Phase A — Architecture tests

- [ ] Every required concept is represented exactly once in the appropriate layer.
- [ ] Product and Container are separate.
- [ ] Barcode and ContainerID uniqueness rules are explicit.
- [ ] Product, batch, and Container attributes are correctly separated.
- [ ] Status meanings do not overlap.
- [ ] All transaction state transitions are defined.
- [ ] Stock-count inclusion/exclusion rules are explicit.
- [ ] Append-only correction behavior is explicit.

## Phase B — Macro-free workbook structural tests

- [ ] Exact worksheet set matches the draft contract.
- [ ] Exact Excel Table names and columns match the draft contract.
- [ ] Primary keys are nonblank and unique in fixtures.
- [ ] Product/Container/Supplier/Location references are valid.
- [ ] Duplicate barcodes are detected.
- [ ] Text barcode formatting preserves leading zeroes.
- [ ] Controlled lists reject invalid values.
- [ ] Formula cells are protected from ordinary editing.
- [ ] Entry cells remain editable.

## Phase C — Formula and business-rule tests

- [ ] Available Container counts are correct.
- [ ] Take/open reduces available count by one.
- [ ] Return restores count where allowed.
- [ ] Disposal removes one qualifying Container.
- [ ] Expired/damaged/missing Containers are excluded.
- [ ] Low-stock, out-of-stock, and reorder outputs are correct.
- [ ] Expired and 30/60/90-day warnings are correct at boundaries.
- [ ] Dashboard totals reconcile to source Tables.
- [ ] Product and Container lookups return expected fixtures.

## Phase D — VBA module tests (future; prohibited in initial task)

- [ ] Barcode lookup
- [ ] Duplicate-scan guard
- [ ] State/transaction validation
- [ ] Atomic transaction append
- [ ] Container update
- [ ] Receiving
- [ ] Scan reset and focus
- [ ] Error recovery and diagnostics
- [ ] Backup/recovery

## Phase E — Integration and physical acceptance (future)

Prove:

`Scan -> identify -> validate -> select transaction -> confirm -> append transaction -> update container -> update stock -> display result -> reset -> scanner focus`

Include normal and abnormal cases, keyboard entry, physical USB scanner, Enter suffix, rapid repeated scans, user cancellation, workbook/file errors, and rollback verification.
