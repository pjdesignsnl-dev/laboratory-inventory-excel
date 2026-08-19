## Label layout check - 2026-08-19 23:01:40

### Code 128 pattern
- Barcode: 0000001
- Pattern length: 10 (Start B + 7 digits + check + Stop)
- Pattern matches stored Barcode value: True (generated from the stored text barcode)

### Label layout (recommended minimum identity)
```
{ Code 128 scannable pattern }  - modCode128.Code128Pattern
0000001                         - human-readable barcode
Product short name              - from tblProducts.ProductName
C000001                         - ContainerID
Batch/Lot (optional)
```

Leading zeros preserved: pattern generated from the 7-digit text 0000001 (no numeric coercion).

### Status
- Code 128 font: not installed on this machine; a licensed Code 128 TrueType font is required on the printer PC (none is distributed through this repository).
- Label printer: not available; printable layout is validated via Excel export to PDF.
- Physical print + real-scanner acceptance: PENDING (owner/hardware dependency).