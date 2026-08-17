# Laboratory Inventory for Excel

A professional, barcode-driven laboratory inventory application implemented in Microsoft Excel for Windows.

## Product goal

Replace a paper-based stock process with a reliable system that tracks each physical bottle, box, package, or container by a unique barcode. The application must provide fast scanner-driven transactions, current container-level stock, expiry and reorder warnings, and a permanent append-only audit history.

This is not a generic inventory template. It is a small inventory-management application built inside Excel.

## Fixed operating context

- Windows desktop
- Microsoft 365 Excel or Excel 2021/2024
- VBA/macros permitted
- One dedicated inventory-editing PC
- One USB keyboard-wedge barcode scanner
- Other computers require read-only information
- No simultaneous inventory editors required
- Smallest tracked unit: one complete physical container/package
- No remaining-volume or individual-piece consumption tracking in v1

## Required development order

1. Requirements and assumptions
2. Workbook architecture and data model
3. Architecture review
4. Macro-free workbook construction
5. Non-VBA architecture and formula testing
6. VBA design against the frozen workbook contract
7. Incremental VBA implementation and module testing
8. Full integration testing
9. Release, deployment, backup, and maintenance documentation

**Do not let VBA drive or silently change the workbook architecture.**

## Current repository state

This initial commit contains the authoritative requirements, operating instructions for an IDE agent, decision and architecture scaffolding, and test/release directories. It deliberately contains no production workbook and no VBA.

Start with [`docs/INITIAL_TASK.md`](docs/INITIAL_TASK.md).

## Main files

- [`AGENTS.md`](AGENTS.md) — non-negotiable rules for all coding agents
- [`docs/requirements.md`](docs/requirements.md) — authoritative product requirements
- [`docs/default-assumptions.md`](docs/default-assumptions.md) — approved defaults for a fast basic v0.1
- [`docs/INITIAL_TASK.md`](docs/INITIAL_TASK.md) — first DeepSeek harness assignment
- [`docs/architecture.md`](docs/architecture.md) — architecture review document to complete
- [`docs/decisions.md`](docs/decisions.md) — append-only design decision log
- [`docs/workbook-contract.md`](docs/workbook-contract.md) — exact workbook/VBA contract, frozen before VBA
- [`tests/test-plan.md`](tests/test-plan.md) — incremental acceptance framework

## Data safety

Only synthetic/anonymized fixtures may be committed. Real laboratory, supplier, employee, or regulated data must remain outside Git unless explicitly approved and appropriately protected.
