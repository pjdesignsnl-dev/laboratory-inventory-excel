# Default assumptions for basic v0.1

These defaults are approved to prevent unnecessary blocking. They remain editable during the architecture/non-VBA phase and must be logged if changed.

1. **Visible language:** English.
2. **Internal naming:** stable English PascalCase column names; `tbl` prefix for Excel Tables; `rng` prefix for named input/display ranges.
3. **Workbook type during first build:** `.xlsx`, deliberately macro-free.
4. **Editing model:** one authoritative editor on the dedicated laboratory PC.
5. **Read-only model:** separate macro-free reporting workbook or published read-only copy after the operational workbook is stable.
6. **Barcode model:** unique internal numeric identifiers rendered as Code 128; scanner sends an Enter suffix.
7. **Barcode storage:** store barcode values as text to preserve leading zeroes and avoid scientific notation.
8. **Container IDs:** generated immutable IDs distinct from the visible barcode.
9. **Transaction IDs:** generated immutable IDs; transaction rows are append-only.
10. **Stock source of truth:** dynamic count of current Container records in qualifying available states; no separately editable stock balance.
11. **Quantity model:** one physical container/package equals one stock unit. No remaining volume, mass, or piece count in v1.
12. **Opened handling:** opening is recorded through `OpenedDate` and transaction history. A returned opened bottle may become available again while retaining its opened history.
13. **Initial user model:** no login/account system. Reserve an Operator/User field and capture the Windows username later when practical.
14. **Expiry thresholds:** configurable 30, 60, and 90 days.
15. **Compatibility:** target desktop Excel 2021/2024 and Microsoft 365; clearly label any 365-only convenience formula.
16. **Sample data:** synthetic products only: pipette tips, laboratory tubes, ethanol, a second solvent, and an expiring reagent.
17. **Compliance boundary:** operational inventory control; not represented as a replacement for an SDS, EHS, GMP/GLP, validated LIMS, or formal hazardous-material system.
18. **Architecture review behavior:** build a strong documented v0.1 using these defaults and stop before VBA for review; do not wait on non-critical questions.
