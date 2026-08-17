# Container status and transaction transition matrix

**Status:** DRAFT

The architecture phase must minimize states, separate current availability from historical opening where appropriate, and classify every transition as:

- **ALLOW** — execute without additional warning;
- **CONFIRM** — show a clear warning and require confirmation;
- **BLOCK** — reject without mutation;
- **N/A** — transaction does not apply.

## Status definitions

_TBD_

## Transaction definitions

_TBD_

## Matrix

| Current status | Receive | Take/Open | Return | Transfer | Dispose | Mark Expired | Mark Damaged | Mark Missing | Adjustment/Reversal |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| TBD | | | | | | | | | |

## Side effects

For every allowed transaction document:

- resulting status;
- location behavior;
- opened/disposal date behavior;
- required fields;
- warning/confirmation behavior;
- transaction snapshot fields;
- stock-count effect;
- reversal/correction behavior.
