VBA Stage 1 test results — Constants + workbook contract runtime validator
Date: 2026-08-18
Excel: 16.0.20228.20190 (x64)
Candidate: workbook/LabInventory_v1.0-candidate.xlsm

Build:
- modConstants.bas imported (contract constants: worksheets, tables, columns,
  named ranges, statuses, transaction types, ID formats, message classes)
- modWorkbookContract.bas imported (runtime contract validator, fail-closed)
- modTestHooks.bas imported (test drivers)
- Excel compiled the VBA project on save (no compile errors)

Tests (driven via Application.Run through COM):
- Test_ContractCheck -> "contract:OK diag=[]"
  => All 9 worksheets, 11 ListObjects, critical columns, 9 named ranges,
     6 statuses, 9 transaction types match the frozen contract.
- Test_ContractDrift -> "drift:before=FAIL after=OK"
  => Deleting named range rngScanInput makes ContractValidate FAIL (fail
     closed); restoring it makes ContractValidate OK again. The drift path is
     proven to work.

Defects found and fixed during Stage 1:
1. Collection-based membership test compared items to keys (never matched);
   switched to item-list membership (values stored as items).
2. Join(parts, "; ") on an empty array raised error 9 (Subscript out of
   range) when there were no issues; added an empty-collection guard.
3. ListRow.Range(1) raised error 9; switched to ListRow.Range.Cells(1,1).
4. CheckName resolved whole-column named ranges (Settings!$A:$A -> 1M cells,
   hang risk); changed to existence-only check via wb.Names(name).

RESULT: PASS
