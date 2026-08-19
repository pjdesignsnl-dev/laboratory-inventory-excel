# Production Acceptance — Laboratory Inventory Excel v1.0.0

Date: 2026-08-19 (Excel 16.0.20228.20190 x64, Windows)
Branch: `feat/non-vba-v0.1` — final source HEAD `adabb53ee3aec334723ec7e8a2083bcbf7563b92`
(production-preparation commits follow on top; see release manifest).

## Overall status

**LabInventory v1.0.0 remains a production candidate.** Software acceptance
passes with no known defects; the following production gate(s) remain open:
**physical barcode-scanner acceptance** and **physical label print** (hardware
not available in the build environment), and the **final deployment location /
permissions** require owner provisioning of the network share.

## Gate results

| Gate | Result |
|---|---|
| Exact state verification (HEAD/SHA/branch/tree/AccessVBOM/Excel/contract) | PASS |
| Release-manifest identity cleanup | PASS (final source HEAD recorded) |
| Deployment architecture chosen (D-023 Accepted, supersedes D-013) | PASS (Option A/B: single writer on private network share + read-only report) |
| Macro security: AccessVBOM | 0 (verified; workbook operates with it disabled) |
| Macro trust model | No signing certificate available → controlled Trusted Location documented (scoped to production folder); no fake trust |
| Mark-of-the-Web handling | Documented (unblock + SHA verify; trusted location exempts the writer) |
| Physical scanner acceptance | **PENDING — no scanner available (not fabricated)** |
| Label print acceptance | Layout validated in PDF; Code 128 font + printer PENDING (documented) |
| Clean production data reset | PASS (`LabInventory_v1.0.0-production.xlsm`, empty operational tables, controlled lists preserved, inert fault module) |
| Initial-data-import safety | PASS (templates + validation rules + QA report template + go-live baseline documented) |
| Backup/recovery drill | PASS (backup → SHA recorded → restore → open/contract/formulas/macros/scan → master unchanged) |
| Read-only access acceptance | PASS (macro-free report, no write path, protected, viewer content OK) |
| Production-location smoke (engine, on a copy) | PASS (import → ReceiveOne → TakeOpen → Return → txn append → reopen persists) |
| Production protection/UI review | PASS (test hooks excluded from production binary; inert fault module documented) |
| Final security/integrity checks | PASS (see below) |
| Final production regression | PASS (full matrix; fault injection on candidate copy only, never the master) |

## Final security / integrity checks

- AccessVBOM = 0 ✓ (verified final state; re-verified after the last build
  temporarily set it to 1 — restored to 0 and the production workbook
  re-verified to open/operate with it disabled)
- VBA project compiles (production binary compiles on save) ✓
- Workbook opens without repair ✓
- No formula errors (dashboard + all sheets) ✓
- Frozen contract validates on open ("Laboratory Inventory ready.") ✓
- Barcode/ID/append-only invariants (proven in earlier acceptance; unchanged) ✓
- No synthetic test data in production workbook (empty operational tables) ✓
- Sheet protection enabled on all sheets; workbook structure protected ✓
- Backup destination valid (LABINV_BACKUP_FOLDER outside master share) ✓
- Read-only artifact macro-free and accessible ✓
- Macro trust as documented (Trusted Location) ✓

## Production artifacts

| Artifact | SHA-256 |
|---|---|
| `workbook/LabInventory_v1.0.0-production.xlsm` | `A4B0A2B1814BAD06DAB5F233CF47873E37FE3AA787C6539DD60AF6597F0300E1` |
| `workbook/LabInventory_v1.0.0-readonly-report.xlsx` | `DF18E9D782D5B581C17A789B831C8FC26FFC423DE0A5A09308DDC81A4313B601` |

## Open gates (owner/hardware dependencies)

1. Physical barcode scanner — connect a USB keyboard-wedge Code 128 scanner and
   run `evidence/production/scanner-acceptance.md` (14-point checklist).
2. Physical label printer + licensed Code 128 font — print and scan
   representative labels.
3. Provision `\\<LAB-SERVER>\Inventory\` share with the D-023 permission model
   (single writer / read-only viewers).
4. Import the actual laboratory inventory via the templates
   (`docs/initial-data-import.md`) and record the go-live baseline.

## Final acceptance language

- **Now:** "LabInventory v1.0.0 remains a production candidate. Software
  acceptance passes with no known defects; the following production gate(s)
  remain open: physical barcode-scanner acceptance, physical label print,
  final deployment share provisioning, initial inventory import."
- **After all gates pass:** "LabInventory v1.0.0 is production ready. All
  defined requirements, automated/runtime acceptance tests, deployment checks,
  backup/recovery checks, and physical scanner acceptance pass with no known
  defects." (Update this file and the release manifest then.)
