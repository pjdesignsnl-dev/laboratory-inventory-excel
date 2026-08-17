# AGENTS.md — Laboratory Inventory Excel

These instructions apply to every agent and every file in this repository.

## Mission

Build a professional, barcode-driven laboratory inventory application in Microsoft Excel for Windows. Track individual physical containers/packages, calculate stock as container counts, provide expiry and reorder controls, and preserve an append-only transaction history.

## Authority order

When instructions conflict, use this order:

1. Explicit current user instruction
2. `docs/requirements.md`
3. An approved decision in `docs/decisions.md`
4. The frozen `docs/workbook-contract.md`
5. This file
6. Other repository documentation

Never silently resolve a conflict. Record the conflict and the chosen resolution in `docs/decisions.md`.

## Non-negotiable architecture rule

Do not write, import, embed, or generate VBA until all of the following are true:

- worksheets are finalized;
- Excel Table names are finalized;
- exact column names and data types are finalized;
- IDs and key relationships are finalized;
- statuses and state transitions are finalized;
- transaction types and effects are finalized;
- controlled lists are finalized;
- named ranges and Scan-screen cell/range locations are finalized;
- formulas and validation are implemented and tested;
- `docs/workbook-contract.md` is marked `FROZEN FOR VBA`;
- the architecture/non-VBA test report passes.

If VBA later exposes an architectural flaw, stop. Document the issue, propose a contract change, and do not rewrite around it silently.

## Development behavior

- Work incrementally and leave the workspace in a reviewable state.
- Use sensible documented assumptions instead of asking broad or avoidable questions.
- Ask only when an unresolved ambiguity could materially affect safety, data integrity, or the user workflow.
- For the basic v0.1, use `docs/default-assumptions.md` unless contradicted by requirements.
- Do not create a giant all-in-one VBA module.
- Do not claim desktop Excel/VBA/scanner tests passed unless they actually ran on Windows desktop Excel.
- Never use Google Sheets as the implementation runtime and never convert the operational workbook to native Google Sheets.
- Preserve Excel 2021/2024 compatibility wherever practical; label Microsoft 365-only functionality explicitly.
- Prefer Excel Tables, structured references, named ranges, constants, and explicit schema validation over fragile row/cell assumptions.
- Avoid `Select`, `Activate`, and hard-coded data row numbers in VBA.
- Treat the transaction table as append-only. Corrections use compensating/reversal transactions.
- Failed transactions must not leave a half-written log or half-updated container state.
- Never commit real sensitive laboratory data. Use synthetic fixtures only.

## Core domain invariants

1. A Product is a general catalogue item.
2. A Container is one uniquely identifiable physical bottle, box, package, or other inventory unit.
3. Every active Container has exactly one unique internal barcode.
4. Multiple Containers may reference the same Product.
5. Stock is the count of Containers whose current state qualifies as available.
6. Remaining mL, grams, individual tips, tubes, filters, gloves, and similar pieces are not tracked in v1.
7. Transactions are permanent historical events and must never be overwritten by normal operation.
8. Current container state is derived/maintained consistently with the accepted transaction.
9. Expired, disposed, damaged, and missing containers cannot be taken through a normal available-stock transaction.
10. Duplicate barcodes are impossible in accepted data.

## Repository conventions

- Architecture and decisions: `docs/`
- Machine-readable draft/frozen contract: `schema/`
- Macro-free and macro-enabled workbook candidates: `workbook/`
- Exported VBA source modules: `vba/`
- Synthetic fixtures and tests: `data/synthetic/`, `tests/`
- Screenshots, logs, and run evidence: `evidence/`
- Approved release packages only: `releases/`
- Repeatable build/test automation: `scripts/`

## Required evidence per phase

### Architecture

- exact worksheet inventory;
- exact Tables and columns;
- key and relationship definitions;
- status and transaction transition matrix;
- barcode rules;
- stock/reorder/expiry formulas;
- edit/view ownership per sheet;
- trade-offs and unresolved risks.

### Workbook construction

- workbook path and file hash;
- worksheet/Table/name inventory exported to text or CSV;
- representative screenshots;
- formula and validation inventory;
- synthetic fixtures;
- non-VBA test results.

### VBA modules

For every module record:

- file/module name;
- exact destination;
- trigger or public entry point;
- worksheets, Tables, columns, named ranges, and status values used;
- dependencies;
- expected mutations;
- test procedure, expected result, and failure cases;
- actual desktop Excel test evidence.

### Release

- workbook version;
- Git commit SHA;
- workbook contract version;
- workbook checksum;
- Excel version/build tested;
- scanner model/configuration tested, when available;
- acceptance result and known limitations.

## First action

Read `docs/requirements.md`, `docs/default-assumptions.md`, and `docs/INITIAL_TASK.md` completely before modifying anything.
