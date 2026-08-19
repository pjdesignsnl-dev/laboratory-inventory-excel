# Scanner & Label Configuration — Laboratory Inventory Excel v1.0.0

## Barcode format (frozen contract)

- Barcodes are **7 digits**, stored as **text** (`\d{7}`), e.g. `0000001`.
- The scanner must be a **USB keyboard-wedge** device with an **Enter suffix**
  so each scan behaves like "type barcode + Enter" into the focused cell.

## Required scanner configuration

| Setting | Required value |
|---|---|
| Symbology | **Code 128** enabled |
| Enter/CR suffix | **Enabled** (append Enter after each scan) |
| Prefix | none (or disabled) |
| Suffix beyond Enter | none |
| Leading zeros | **Preserved** (barcode `0000001` must transmit as `0000001`) |

Recommended scanner families: Zebra (e.g., LI/DS series), Honeywell
(e.g., Voyager/Genesis), Datalogic (e.g., QuickScan) — any USB keyboard-wedge
Code 128 device configured as above.

## Scanner acceptance status

- **Physical scanner: NOT available in the build/test environment.** Scanner
  behavior was validated by keyboard-wedge simulation (typing the barcode +
  Enter into `Scan!D7` with events enabled). **Real-hardware acceptance is an
  open production gate** — perform it with the actual device before declaring
  production ready. Use the checklist in
  `evidence/production/scanner-acceptance.md` (template; fill in make/model/
  config and results).
- On-site acceptance must cover: known/unknown/invalid barcodes, repeated fast
  scans, scan → Take/Open, scan → Return, scan → Transfer, scan → Dispose
  confirm, expired-by-date TakeOpen blocked, success/error path input reset,
  focus returns to the field, next scan accepted immediately, leading zeros
  preserved.

## Label printing

- `modCode128.Code128Pattern` generates the Code 128B pattern string for a
  barcode (with check digit). Render it with a **TrueType Code 128 font** on
  the label.
- **Required font:** a licensed Code 128 TrueType font on the printer machine.
  None is installed in the build environment and none is distributed through
  this repository (licensing). Install an approved font on the lab printer PC.
- **Label layout (recommended minimum):**
  - Code 128 pattern (scannable) for the Barcode,
  - human-readable Barcode underneath,
  - Product name or short name,
  - ContainerID,
  - optional Batch/Lot.
  Do not overcrowd — see the layout sample in
  `docs/operator-quick-start.md` (labels).
- **Printer status:** no physical label printer was available in the build
  environment; printable output is validated in Excel/PDF (see
  `evidence/production/label-layout-check.md`). Physical print acceptance is
  pending a real printer + font.

## Trusted Location setup (macro security — required)

The production workbook must run macros WITHOUT "enable all macros" and
without lowering Trust Center security. On the **dedicated writer PC only**:

1. Excel → File → Options → Trust Center → Trust Center Settings →
   **Trusted Locations**.
2. Add the production share folder (`\\<LAB-SERVER>\Inventory\`) as a Trusted
   Location. Check **"Allow files on a network location to be trusted"**.
3. Do NOT add broad locations (e.g., whole drives or OneDrive) — scope the
   Trusted Location to the production master folder only.
4. Verify: open the master from the share; the status bar shows "Laboratory
   Inventory ready." and scans work.

Mark-of-the-Web (MOTW): files copied from the repository or a download may be
flagged by Windows. If a copy of the master is blocked ("mark of the web"),
open it once from the Trusted Location or unblock via file properties
(Properties → Unblock) after verifying the SHA-256 matches the release
manifest. The trusted-location master itself is not subject to MOTW on the
writer PC.
