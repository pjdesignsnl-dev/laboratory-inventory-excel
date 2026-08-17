# Workbook architecture

**Status:** DRAFT — not approved for VBA

This document must be completed during the architecture phase and reviewed against `docs/requirements.md`.

## 1. Design summary

_TBD_

## 2. Workbook topology

| Worksheet | Purpose | Editors | Viewers | Excel Tables | Protected areas |
|---|---|---|---|---|---|
| TBD | | | | | |

## 3. Table definitions

For every Table provide exact Table name, worksheet, primary key, columns, data types, mandatory/optional rule, default, validation, uniqueness rule, and description.

_TBD_

## 4. Relationships

_TBD_

## 5. Controlled lists and constants

_TBD_

## 6. Product vs batch vs container data placement

_TBD_

## 7. Barcode design

_TBD_

## 8. Status and transaction model

See `docs/status-transition-matrix.md`.

## 9. Stock, expiry, and reorder logic

_TBD_

## 10. Scan workflow

_TBD_

## 11. Receiving workflow

_TBD_

## 12. Dashboard and reporting

_TBD_

## 13. Protection and manual-edit controls

_TBD_

## 14. Read-only and deployment architecture

_TBD_

## 15. Compatibility

_TBD_

## 16. Trade-offs, risks, and deferred scope

_TBD_

## 17. Architecture review checklist

- [ ] Every worksheet is defined.
- [ ] Every Excel Table has an exact stable name.
- [ ] Every column has an exact stable name and data type.
- [ ] Mandatory/optional rules are explicit.
- [ ] Keys, uniqueness, and relationships are explicit.
- [ ] Statuses and transactions have unambiguous semantics.
- [ ] Every state transition is classified allowed/warning/blocked.
- [ ] Barcode uniqueness and duplicate-scan behavior are defined.
- [ ] Stock is not manually editable.
- [ ] Formula compatibility is documented.
- [ ] Scan and receiving interfaces have exact ranges/named ranges.
- [ ] Protection and recovery limitations are documented.
- [ ] Non-VBA tests are defined.
