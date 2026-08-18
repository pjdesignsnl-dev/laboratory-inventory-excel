Attribute VB_Name = "modTestHooks"
Option Explicit

' ============================================================================
' modTestHooks - Test hooks for VBA-focused tests (extended)
' ============================================================================
' Each Test_* procedure runs a focused workflow and appends PASS/FAIL lines to
' the test log file. All operations are error-suppressed so a failure cannot
' block the COM caller. Not part of the operator workflow.
' ============================================================================

Private Const TEST_LOG As String = "C:\Users\Q\Documents\laboratory-inventory-excel\evidence\vba\test-hooks-output.txt"

Private Sub LogLine(ByVal s As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open TEST_LOG For Append As #ff
    Print #ff, s
    Close #ff
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 1: contract check
Public Sub Test_ContractCheck()
    On Error Resume Next
    Dim ok As Boolean
    ok = modWorkbookContract.ContractValidate(True)
    LogLine "contract:" & IIf(ok, "OK", "FAIL") & " diag=[" & modWorkbookContract.ContractDiagnostics() & "]"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 1: drift (named-range delete)
Public Sub Test_ContractDrift()
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim savedRefersTo As String
    savedRefersTo = wb.Names(RNG_SCAN_INPUT).RefersTo
    wb.Names(RNG_SCAN_INPUT).Delete
    Dim okBefore As Boolean
    okBefore = modWorkbookContract.ContractValidate(True)
    wb.Names.Add name:=RNG_SCAN_INPUT, RefersTo:=savedRefersTo
    Dim okAfter As Boolean
    okAfter = modWorkbookContract.ContractValidate(True)
    LogLine "drift:before=" & IIf(okBefore, "OK", "FAIL") & " after=" & IIf(okAfter, "OK", "FAIL")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 2: barcode lookup (instrumented)
Public Sub Test_BarcodeLookup()
    On Error Resume Next
    LogLine "lookup-start"
    Dim rowNum As Long

    ' --- pure-VBA loop lookup sanity (no worksheet functions) ---
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    Dim rng As Range
    Set rng = lo.ListColumns("Barcode").DataBodyRange
    LogLine "lookup-range:" & rng.Address & " rows=" & rng.Rows.count
    LogLine "lookup-type:" & TypeName(rng.Cells(1, 1).Value2) & " text=[" & rng.Cells(1, 1).text & "]"

    Dim found As Long
    found = 0
    Dim cell As Range
    For Each cell In rng.Cells
        If CStr(cell.Value2) = "0000001" Then
            found = cell.Row - rng.Row + 1
            Exit For
        End If
    Next cell
    LogLine "vbaloop-found:" & found

    ' --- full LookupState + FindBarcodeRow (array-backed, no worksheet fns) ---
    LogLine "lookupbybarcode-start"
    Dim r1 As Long
    r1 = modBarcodeLookup.LookupState("0000001")
    Dim row1 As Long
    row1 = modBarcodeLookup.FindBarcodeRow("0000001")
    LogLine "lookupbybarcode-end:" & IIf(r1 = modBarcodeLookup.LR_FOUND, "OK", "FAIL") & " row=" & row1

    Dim r2 As Long
    r2 = modBarcodeLookup.LookupState("9999999")
    LogLine "unknown-end:" & IIf(r2 = modBarcodeLookup.LR_UNKNOWN, "OK", "FAIL")

    Dim r3 As Long
    r3 = modBarcodeLookup.LookupState("")
    LogLine "lookup-empty:" & IIf(r3 = modBarcodeLookup.LR_EMPTY, "OK", "FAIL")

    Dim r4 As Long
    r4 = modBarcodeLookup.LookupState("abc1234")
    LogLine "lookup-invalid:" & IIf(r4 = modBarcodeLookup.LR_INVALID, "OK", "FAIL")
    LogLine "lookup-end"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 3: transition validation
Public Sub Test_Transitions()
    On Error Resume Next
    Dim msg As String
    Dim mc As msgClass
    ' Available -> TakeOpen allowed
    LogLine "t-avail-takeopen:" & IIf(modValidation.ValidateTransition(STATUS_AVAILABLE, TXN_TAKE_OPEN, Empty, msg, mc), "OK", "FAIL")
    ' Available expired-by-date -> TakeOpen must be BLOCKED (D-018)
    LogLine "t-expiredby-date-takeopen:" & IIf(Not modValidation.ValidateTransition(STATUS_AVAILABLE, TXN_TAKE_OPEN, Date - 1, msg, mc), "OK", "FAIL")
    ' InUse -> Return allowed
    LogLine "t-inuse-return:" & IIf(modValidation.ValidateTransition(STATUS_IN_USE, TXN_RETURN, Empty, msg, mc), "OK", "FAIL")
    ' Disposed -> TakeOpen blocked
    LogLine "t-disposed-takeopen:" & IIf(Not modValidation.ValidateTransition(STATUS_DISPOSED, TXN_TAKE_OPEN, Empty, msg, mc), "OK", "FAIL")
    ' Expired -> Dispose allowed
    LogLine "t-expired-dispose:" & IIf(modValidation.ValidateTransition(STATUS_EXPIRED, TXN_DISPOSE, Empty, msg, mc), "OK", "FAIL")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 4: transaction ID + snapshot
Public Sub Test_TransactionID()
    On Error Resume Next
    Dim tid As String
    tid = modTransactions.NextTransactionID()
    LogLine "txnid-format:" & IIf(Len(tid) = 9 And Left$(tid, 1) = "T", "OK", "FAIL") & " tid=" & tid
    LogLine "txnid-unique:" & IIf(Not modTransactions.TransactionIDExists(tid), "OK", "FAIL")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 7/8: receive one + receive N
Public Sub Test_ReceiveOne()
    On Error Resume Next
    Dim cid As String
    cid = modReceiving.ReceiveOne("P000001", "LOT-T1", Empty, Empty, "LOC0004", "VBA test receive 1")
    LogLine "receive-one:" & IIf(Len(cid) > 0, "OK", "FAIL") & " cid=" & cid
    ' verify container + transaction exist
    LogLine "receive-one-cid:" & IIf(modContainers.ContainerIDExists(cid), "OK", "FAIL")
    On Error GoTo 0
End Sub

Public Sub Test_ReceiveN()
    On Error Resume Next
    Dim res() As String
    res = modReceiving.ReceiveN("P000001", "LOT-TN", Empty, Empty, "LOC0004", 3, "VBA test receive N")
    LogLine "receive-n:" & IIf(UBound(res) = 3, "OK", "FAIL") & " count=" & UBound(res)
    LogLine "receive-n-ids:" & res(1) & "," & res(2) & "," & res(3)
    LogLine "receive-n-unique:" & IIf(res(1) <> res(2) And res(2) <> res(3) And res(1) <> res(3), "OK", "FAIL")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 5/6/9: atomic commit via scan interface (diagnosed)
Public Sub Test_CommitTakeOpen()
    ' Scan barcode 0000001, then commit TakeOpen.
    On Error Resume Next
    modScanInterface.HandleScannedBarcode "0000001"
    ' read staging barcode for diagnostics
    Dim staging As ListObject
    Set staging = ThisWorkbook.Worksheets("Scan").ListObjects("tblScanResults")
    Dim sb As String
    sb = CStr(staging.DataBodyRange.Cells(1, modBarcodeLookup.ColumnIndex(staging, "Barcode")).Value2)
    Dim ls As Long
    ls = modBarcodeLookup.LookupState(sb)
    LogLine "commit-diag:staging-barcode=[" & sb & "] lookupstate=" & ls

    modScanInterface.CommitAction TXN_TAKE_OPEN
    ' check status message
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    Dim m As String
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    LogLine "commit-takeopen:" & IIf(InStr(m, "OK") > 0, "OK", "FAIL") & " msg=" & m
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 13: Code128
Public Sub Test_Code128()
    On Error Resume Next
    Dim pat As String
    pat = modCode128.Code128Pattern("0000001")
    LogLine "code128:" & IIf(Len(pat) >= 9, "OK", "FAIL") & " len=" & Len(pat)
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 10: operator
Public Sub Test_Operator()
    On Error Resume Next
    Dim op As String
    op = modUtilities.GetOperator()
    LogLine "operator:" & IIf(Len(op) > 0, "OK", "FAIL") & " op=" & op
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 11: backup
Public Sub Test_Backup()
    On Error Resume Next
    Dim path As String
    path = modBackup.CreateBackup("test backup")
    LogLine "backup:" & IIf(Len(path) > 0, "OK", "FAIL") & " path=" & path
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ full sweep
Public Sub Test_RunAll()
    Test_ContractCheck
    Test_ContractDrift
    Test_BarcodeLookup
    Test_Transitions
    Test_TransactionID
    Test_ReceiveOne
    Test_ReceiveN
    Test_Code128
    Test_Operator
    Test_Backup
End Sub

' ------------------------------------------------------------------ core sweep (no commit/backup)
Public Sub Test_RunCore()
    Test_ContractCheck
    Test_BarcodeLookup
    Test_Transitions
    Test_TransactionID
    Test_ReceiveOne
    Test_ReceiveN
    Test_Code128
    Test_Operator
End Sub

' ------------------------------------------------------------------ minimal (contract only) - used to isolate the self-test path
Public Sub Test_Minimal()
    Test_ContractCheck
End Sub
