# Initial DeepSeek harness task — architecture and macro-free v0.1

Read `AGENTS.md`, `docs/requirements.md`, and `docs/default-assumptions.md` completely. Treat them as authoritative.

## Goal

Create a professional but still editable **basic v0.1 macro-free workbook** and the documentation needed to review it. Use sensible assumptions from `docs/default-assumptions.md`; do not block on non-critical questions.

## Authorized scope

Complete Stages 1 through 5 only:

1. Requirements and assumptions analysis
2. Complete workbook architecture
3. Architecture review documentation
4. Macro-free workbook construction
5. Non-VBA architecture/formula/validation testing

**VBA is explicitly forbidden in this task.** Do not create `.bas`, `.cls`, `.frm`, embedded macros, Office Scripts, or another automation layer.

## Required work

1. Inspect the Windows/Excel environment and record the actual Excel edition/build if discoverable without changing system configuration.
2. Produce `docs/requirements-analysis.md` covering assumptions, ambiguities, risks, and decisions.
3. Complete `docs/architecture.md` with:
   - worksheets and purpose;
   - editors/viewers;
   - exact Excel Tables;
   - exact columns, data types, mandatory/optional rules;
   - keys and relationships;
   - controlled lists;
   - barcode rules;
   - stock, reorder, expiry, and dashboard calculation logic;
   - product/batch/container field placement;
   - scan and receiving workflows;
   - state-transition rules;
   - read-only deployment recommendation and trade-offs.
4. Complete `docs/status-transition-matrix.md`.
5. Update `docs/decisions.md` with every material design choice.
6. Populate `schema/workbook-contract.yaml` as `status: draft`.
7. Construct `workbook/LabInventory_v0.1.xlsx` using desktop Excel or an Excel-compatible generation method, then verify it in desktop Excel.
8. Include the required sheets, final proposed Tables, formulas, validation, conditional formatting, a practical Dashboard, a polished Scan interface, a receiving interface/foundation, protection design, and synthetic worked examples.
9. Make the workbook function as far as reasonably possible without VBA.
10. Create automated or repeatable structural inspections where practical and record all manual tests.
11. Complete `tests/non-vba-results.md` with actual results and evidence paths.
12. Export a text inventory of worksheets, Tables, columns, named ranges, formulas, validation rules, and protected areas into `evidence/`.
13. Capture representative screenshots in `evidence/screenshots/`.
14. Commit the work on a dedicated branch such as `feat/non-vba-v0.1`.

## Required stop condition

Stop after presenting:

- architecture summary and trade-offs;
- exact workbook path;
- workbook SHA-256;
- Git branch, commit SHA, and clean/dirty state;
- non-VBA test summary;
- evidence inventory;
- unresolved decisions that genuinely need owner input;
- explicit statement: `NO VBA HAS BEEN WRITTEN OR EMBEDDED`.

Do not mark `docs/workbook-contract.md` as frozen and do not proceed to VBA until the owner explicitly approves the architecture and non-VBA workbook.
