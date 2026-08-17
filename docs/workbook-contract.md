# Frozen workbook contract

**Status:** NOT FROZEN — VBA PROHIBITED

This document becomes the human-readable binding contract between the workbook and VBA. Populate it after architecture review and non-VBA construction. Change the status to `FROZEN FOR VBA` only after explicit owner approval.

## Workbook identity

- Product name: TBD
- Workbook version: TBD
- Contract version: TBD
- Target Excel versions: TBD
- Architecture approval reference: TBD

## Worksheets

_TBD_

## Excel Tables and exact columns

_TBD_

## Named ranges and exact cells

_TBD_

## IDs and generation rules

_TBD_

## Controlled status values

_TBD_

## Controlled transaction types

_TBD_

## Allowed state transitions

See `docs/status-transition-matrix.md`.

## Formula contract

_TBD_

## Validation and protection contract

_TBD_

## VBA dependency declaration

No VBA exists yet. When authorized, every module must list the exact contract members it uses.

## Change control

After freezing, any change to a worksheet, Table, column, named range, status, transaction type, formula contract, or Scan interface location requires:

1. documented problem;
2. proposed contract revision;
3. impact analysis;
4. owner approval;
5. contract version increase;
6. affected VBA and test updates;
7. full regression test.
