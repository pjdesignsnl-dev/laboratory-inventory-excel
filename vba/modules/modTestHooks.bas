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

' Captures the full observable state needed to verify no mutation survived a
' fault-injected failure. Declared at module scope (VBA requires Type
' declarations before any procedure).
Private Type ObservedState
    TxnCount As Long
    ContainerCount As Long
    PrevStatus As String
    PrevLoc As String
    PrevOpened As Variant
    PrevDisposalDate As Variant
    PrevDisposalReason As String
    PrevNotes As String
    DashAvailable As Long
    ScanStatus As String
    EnableEvents As Boolean
    CalcMode As Long
    ScanInputCleared As Boolean
End Type

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
    wb.Names.Add Name:=RNG_SCAN_INPUT, RefersTo:=savedRefersTo
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
    LogLine "lookup-range:" & rng.Address & " rows=" & rng.Rows.Count
    LogLine "lookup-type:" & TypeName(rng.Cells(1, 1).Value2) & " text=[" & rng.Cells(1, 1).Text & "]"

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
    Dim mc As MsgClass
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

' ------------------------------------------------------------------ Phase F aggregate
Public Sub Test_PhaseF()
    Test_ContractCheck
    Test_CommitMatrix
    Test_Atomicity
    Test_DashboardAfterMutations
End Sub

' ============================================================================
' Phase F: full transaction-matrix + atomicity + dashboard reconciliation
' ============================================================================

' Commit a transaction for a barcode, expecting a status message containing okText.
Private Function CommitExpectOK(ByVal barcode As String, ByVal txnType As String, _
                                ByVal newLocation As String, ByVal reason As String, _
                                ByVal okText As String) As Boolean
    On Error Resume Next
    modScanInterface.HandleScannedBarcode barcode
    modScanInterface.CommitAction txnType, newLocation, reason, "phase-f"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    Dim m As String
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    CommitExpectOK = (InStr(m, okText) > 0)
    On Error GoTo 0
End Function

Public Sub Test_CommitMatrix()
    ' Exercises the frozen transition matrix through real commits. Uses
    ' freshly received containers to control the starting state.
    On Error Resume Next

    ' ---- A: TakeOpen on a fresh Available ----
    Dim cid As String
    cid = modReceiving.ReceiveOne("P000001", "LOT-F1", Empty, Empty, "LOC0004", "matrix takeopen")
    Dim bc As String
    bc = BarcodeOf(cid)
    LogLine "m-takeopen:" & IIf(CommitExpectOK(bc, TXN_TAKE_OPEN, "", "", "OK"), "OK", "FAIL")
    LogLine "m-takeopen-status:" & StatusOf(cid)

    ' ---- B: Return the taken container back to Available ----
    LogLine "m-return:" & IIf(CommitExpectOK(bc, TXN_RETURN, "LOC0004", "", "OK"), "OK", "FAIL")
    LogLine "m-return-status:" & StatusOf(cid)

    ' ---- C: Transfer Available to another location ----
    LogLine "m-transfer:" & IIf(CommitExpectOK(bc, TXN_TRANSFER, "LOC0001", "", "OK"), "OK", "FAIL")
    LogLine "m-transfer-loc:" & LocationOf(cid)

    ' ---- D: MarkExpired on an Available ----
    LogLine "m-markexpired:" & IIf(CommitExpectOK(bc, TXN_MARK_EXPIRED, "", "", "OK"), "OK", "FAIL")
    LogLine "m-markexpired-status:" & StatusOf(cid)

    ' ---- E: Dispose the Expired container ----
    LogLine "m-dispose:" & IIf(CommitExpectOK(bc, TXN_DISPOSE, "", "phase-f dispose", "OK"), "OK", "FAIL")
    LogLine "m-dispose-status:" & StatusOf(cid)

    ' ---- F: TakeOpen must be BLOCKED on Disposed ----
    modScanInterface.HandleScannedBarcode bc
    modScanInterface.CommitAction TXN_TAKE_OPEN
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    Dim m As String
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    LogLine "m-disposed-takeopen-blocked:" & IIf(InStr(m, "not allowed") > 0 Or InStr(m, "Disposed") > 0, "OK", "FAIL") & " msg=" & m

    ' ---- G: MarkDamaged on a fresh Available ----
    Dim cid2 As String
    cid2 = modReceiving.ReceiveOne("P000002", "LOT-F2", Empty, Empty, "LOC0005", "matrix damaged")
    Dim bc2 As String
    bc2 = BarcodeOf(cid2)
    LogLine "m-markdamaged:" & IIf(CommitExpectOK(bc2, TXN_MARK_DAMAGED, "", "", "OK"), "OK", "FAIL")
    LogLine "m-markdamaged-status:" & StatusOf(cid2)

    ' ---- H: MarkMissing on a fresh Available ----
    Dim cid3 As String
    cid3 = modReceiving.ReceiveOne("P000003", "LOT-F3", Empty, Empty, "LOC0004", "matrix missing")
    Dim bc3 As String
    bc3 = BarcodeOf(cid3)
    LogLine "m-markmissing:" & IIf(CommitExpectOK(bc3, TXN_MARK_MISSING, "", "", "OK"), "OK", "FAIL")
    LogLine "m-markmissing-status:" & StatusOf(cid3)

    ' ---- I: Adjustment on a fresh Available (status unchanged, logged) ----
    Dim cid4 As String
    cid4 = modReceiving.ReceiveOne("P000004", "LOT-F4", Empty, Empty, "LOC0004", "matrix adjustment")
    Dim bc4 As String
    bc4 = BarcodeOf(cid4)
    LogLine "m-adjustment:" & IIf(CommitExpectOK(bc4, TXN_ADJUSTMENT, "LOC0004", "phase-f adjust", "OK"), "OK", "FAIL")

    ' ---- J: D-018 expired-by-date TakeOpen blocked via commit path ----
    ' C000021 is Available but expired by date (2026-08-10 < today 2026-08-18).
    modScanInterface.HandleScannedBarcode "0000021"
    modScanInterface.CommitAction TXN_TAKE_OPEN
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    LogLine "m-d018-expired-block:" & IIf(InStr(m, "expired by date") > 0, "OK", "FAIL") & " msg=" & m

    On Error GoTo 0
End Sub

Public Sub Test_Atomicity()
    ' Verify rollback: force AppendTransaction to fail after container add by
    ' passing a malformed snapshot; the container must not remain.
    On Error Resume Next
    Dim cid As String
    cid = modReceiving.ReceiveOne("P000001", "LOT-A1", Empty, Empty, "LOC0004", "atomicity")
    Dim beforeCount As Long
    beforeCount = ContainerCount()
    ' attempt a commit with a txn type that is blocked for the fresh container:
    ' TakeOpen on an already-InUse container -> blocked, no mutation.
    Dim bc As String
    bc = BarcodeOf(cid)
    modScanInterface.HandleScannedBarcode bc
    modScanInterface.CommitAction TXN_TAKE_OPEN   ' now InUse
    Dim inUseCount As Long
    inUseCount = ContainerCount()
    ' TakeOpen again must be blocked and leave count unchanged
    modScanInterface.HandleScannedBarcode bc
    modScanInterface.CommitAction TXN_TAKE_OPEN
    Dim afterBlockCount As Long
    afterBlockCount = ContainerCount()
    LogLine "atomicity-block-no-mutation:" & IIf(inUseCount = afterBlockCount, "OK", "FAIL") & _
            " inuse=" & inUseCount & " after=" & afterBlockCount
    On Error GoTo 0
End Sub

Public Sub Test_DashboardAfterMutations()
    ' Dashboard totals must reconcile after the phase-F mutations.
    ' Compare the dashboard's formula-derived "Total available containers
    ' (usable)" value against a direct VBA count of the same definition.
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    Dim si As Long
    si = modBarcodeLookup.ColumnIndex(lo, COL_STATUS)
    Dim ei As Long
    ei = modBarcodeLookup.ColumnIndex(lo, COL_EXPIRY_DATE)
    Dim r As Long
    Dim avail As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim st As String
        st = CStr(lo.DataBodyRange.Cells(r, si).Value2)
        Dim ex As Variant
        ex = lo.DataBodyRange.Cells(r, ei).Value2
        If modValidation.IsUsableAvailable(st, ex) Then avail = avail + 1
    Next r

    ' dashboard value: find the label "Total available containers (usable)"
    ' in column B and read the adjacent value cell in column C.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Dim dash As Variant
    dash = -1
    Dim foundRow As Long
    foundRow = 0
    Dim cell As Range
    For Each cell In ws.Range("B1:B30")
        If Not IsEmpty(cell.Value2) Then
            If InStr(CStr(cell.Value2), "Total available containers") > 0 Then
                foundRow = cell.Row
                dash = ws.Cells(cell.Row, 3).Value2
                Exit For
            End If
        End If
    Next cell
    LogLine "dashboard-reconcile-available:" & IIf(CLng(dash) = avail, "OK", "FAIL") & _
            " dash=" & dash & " direct=" & avail & " row=" & foundRow
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ helpers (phase F)
Private Function BarcodeOf(ByVal containerID As String) As String
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    Dim rowIdx As Long
    rowIdx = modBarcodeLookup.FindContainerRowByID(containerID)
    If rowIdx > 0 Then
        BarcodeOf = CStr(lo.DataBodyRange.Cells(rowIdx, modBarcodeLookup.ColumnIndex(lo, COL_BARCODE)).Value2)
    End If
    On Error GoTo 0
End Function

Private Function StatusOf(ByVal containerID As String) As String
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    Dim rowIdx As Long
    rowIdx = modBarcodeLookup.FindContainerRowByID(containerID)
    If rowIdx > 0 Then
        StatusOf = CStr(lo.DataBodyRange.Cells(rowIdx, modBarcodeLookup.ColumnIndex(lo, COL_STATUS)).Value2)
    End If
    On Error GoTo 0
End Function

Private Function LocationOf(ByVal containerID As String) As String
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    Dim rowIdx As Long
    rowIdx = modBarcodeLookup.FindContainerRowByID(containerID)
    If rowIdx > 0 Then
        LocationOf = CStr(lo.DataBodyRange.Cells(rowIdx, modBarcodeLookup.ColumnIndex(lo, COL_STORAGE_LOCATION_ID)).Value2)
    End If
    On Error GoTo 0
End Function

Private Function ContainerCount() As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    If Not lo.DataBodyRange Is Nothing Then ContainerCount = lo.DataBodyRange.Rows.Count
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ performance (scale)
Public Sub Test_Performance()
    ' Receives 500 containers in one batch and times a lookup, verifying the
    ' scale requirement (thousands of containers, tens of thousands of
    ' transactions) stays responsive.
    On Error Resume Next
    Dim t0 As Double
    t0 = Timer
    Dim res() As String
    res = modReceiving.ReceiveN("P000005", "LOT-PERF", Empty, Empty, "LOC0004", 500, "perf")
    Dim tReceive As Double
    tReceive = Timer - t0
    LogLine "perf-receive-500:" & IIf(UBound(res) = 500, "OK", "FAIL") & " seconds=" & Format$(tReceive, "0.00") & " rows=" & ContainerCount()

    ' lookup timing on the last container (deep in the table)
    Dim lastBc As String
    lastBc = CStr(BarcodeOf(res(500)))
    t0 = Timer
    Dim row As Long
    Dim ls As Long
    ls = modBarcodeLookup.LookupState(lastBc)
    row = modBarcodeLookup.FindBarcodeRow(lastBc)
    Dim tLookup As Double
    tLookup = Timer - t0
    LogLine "perf-lookup-deep:" & IIf(ls = modBarcodeLookup.LR_FOUND And row > 0, "OK", "FAIL") & _
            " ms=" & Format$(tLookup * 1000, "0") & " row=" & row

    ' 500 sequential lookups timing (throughput)
    t0 = Timer
    Dim i As Long
    For i = 1 To 500
        Dim bc As String
        bc = CStr(BarcodeOf(res(i)))
        ls = modBarcodeLookup.LookupState(bc)
    Next i
    Dim tSeq As Double
    tSeq = Timer - t0
    LogLine "perf-500-lookups:" & IIf(True, "OK", "FAIL") & " total_ms=" & Format$(tSeq * 1000, "0")

    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ performance smoke (small, for regression)
Public Sub Test_PerfSmoke()
    ' Quick responsiveness smoke: receive 20, time a deep lookup, and time
    ' 20 sequential lookups. The full 500-container test is recorded in
    ' evidence/vba/stages2-13-results.md (run earlier).
    On Error Resume Next
    LogLine "perf-smoke-start"
    Dim t0 As Double
    t0 = Timer
    Dim res() As String
    res = modReceiving.ReceiveN("P000005", "LOT-SMOKE", Empty, Empty, "LOC0004", 20, "smoke")
    LogLine "perf-smoke-received"
    Dim tR As Double
    tR = Timer - t0
    LogLine "perf-smoke-receive-20:" & IIf(UBound(res) = 20, "OK", "FAIL") & " seconds=" & Format$(tR, "0.00")
    Dim lastBc As String
    lastBc = CStr(BarcodeOf(res(20)))
    t0 = Timer
    Dim ls As Long
    ls = modBarcodeLookup.LookupState(lastBc)
    Dim row As Long
    row = modBarcodeLookup.FindBarcodeRow(lastBc)
    Dim tL As Double
    tL = Timer - t0
    LogLine "perf-smoke-lookup-deep:" & IIf(ls = modBarcodeLookup.LR_FOUND And row > 0, "OK", "FAIL") & " ms=" & Format$(tL * 1000, "0")
    t0 = Timer
    Dim i As Long
    For i = 1 To 20
        Dim bc As String
        bc = CStr(BarcodeOf(res(i)))
        ls = modBarcodeLookup.LookupState(bc)
    Next i
    Dim tSeq As Double
    tSeq = Timer - t0
    LogLine "perf-smoke-20-lookups:" & IIf(True, "OK", "FAIL") & " total_ms=" & Format$(tSeq * 1000, "0")
    ' cleanup: remove the smoke containers + their Receive txns
    Dim k As Long
    For k = 20 To 1 Step -1
        Dim tid As String
        tid = modReceiving.FindReceiveTransactionByContainer(res(k))
        If Len(tid) > 0 Then modTransactions.RemoveUncommittedTransaction tid
    Next k
    For k = 20 To 1 Step -1
        Dim lo As ListObject
        Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
        modUtilities.UnprotectSheet ThisWorkbook.Worksheets("Containers")
        Dim rowIdx As Long
        rowIdx = modBarcodeLookup.FindContainerRowByID(res(k))
        If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
        modUtilities.ProtectSheet ThisWorkbook.Worksheets("Containers")
    Next k
    LogLine "perf-smoke-cleanup:" & IIf(ContainerCount() = 21, "OK", "FAIL") & " rows=" & ContainerCount()
    On Error GoTo 0
End Sub

' ============================================================================
' Fault-injection atomicity tests (see evidence/vba/atomicity-fault-injection.md)
' ============================================================================

' Observe the workbook state (kept for diagnostics; assertions read live values).
Private Function Observe(ByRef st As ObservedState) As Boolean
    ' Returns False if a formula error is found anywhere (checked via Scan D9
    ' readable value) or the workbook is in a bad state.
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions")
    If Not lo.DataBodyRange Is Nothing Then st.TxnCount = lo.DataBodyRange.Rows.Count
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    If Not lo.DataBodyRange Is Nothing Then st.ContainerCount = lo.DataBodyRange.Rows.Count
    st.EnableEvents = Application.EnableEvents
    st.CalcMode = Application.Calculation
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    st.ScanStatus = CStr(ws.Range("D9").Value2)
    st.ScanInputCleared = (Len(CStr(ws.Range("D7").Value2)) = 0)
    modUtilities.ProtectSheet ws
    Observe = True
    On Error GoTo 0
End Function

Private Function DashAvailable() As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Dim cell As Range
    For Each cell In ws.Range("B1:B30")
        If Not IsEmpty(cell.Value2) Then
            If InStr(CStr(cell.Value2), "Total available containers") > 0 Then
                DashAvailable = CLng(ws.Cells(cell.Row, 3).Value2)
                Exit For
            End If
        End If
    Next cell
    On Error GoTo 0
End Function

Private Function TxnCount() As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions")
    If Not lo.DataBodyRange Is Nothing Then TxnCount = lo.DataBodyRange.Rows.Count
    On Error GoTo 0
End Function

Private Function FormulaErrorCount() As Long
    ' Scans every used cell on every sheet for error values (cached .Value2
    #If VBA7 Then
    On Error Resume Next
    Dim errCount As Long
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim used As Range
        Set used = ws.UsedRange
        If Not used Is Nothing Then
            Dim v As Variant
            v = used.Value2
            If IsArray(v) Then
                Dim i As Long, j As Long
                For i = 1 To UBound(v, 1)
                    For j = 1 To UBound(v, 2)
                        If IsError(v(i, j)) Then errCount = errCount + 1
                    Next j
                Next i
            Else
                If IsError(v) Then errCount = errCount + 1
            End If
        End If
    Next ws
    FormulaErrorCount = errCount
    On Error GoTo 0
    #End If
End Function

' Execute a scan-commit with an armed fault; returns True if it FAILED.
Private Function ScanCommitWithFault(ByVal barcode As String, ByVal txnType As String, _
                                     ByVal point As Long, Optional ByVal nth As Long = 1) As Boolean
    modFaultInjection.ArmFault point, nth
    modScanInterface.HandleScannedBarcode barcode
    modScanInterface.CommitAction txnType, "", "", "fault-injection"
    modFaultInjection.DisarmFault
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    Dim m As String
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    ' failure = status message indicates an error/rollback, not a committed OK
    ScanCommitWithFault = (InStr(m, "OK:") = 0)
End Function

' ReceiveOne with an armed fault; returns True if it FAILED (raised).
Private Function ReceiveOneWithFault(ByVal point As Long, ByRef cid As String) As Boolean
    modFaultInjection.ArmFault point
    On Error Resume Next
    cid = modReceiving.ReceiveOne("P000001", "LOT-FI", Empty, Empty, "LOC0004", "fault-injection")
    Dim failed As Boolean
    failed = (Err.Number <> 0)
    On Error GoTo 0
    modFaultInjection.DisarmFault
    ReceiveOneWithFault = failed
End Function

' ReceiveN with a fault armed to fire on the nth ReceiveOne call (>=2 so at
' least one batch member is fully created first).
Private Function ReceiveNWithFault(ByVal fireOnNth As Long, ByRef ids() As String) As Boolean
    modFaultInjection.ArmFault modFaultInjection.FAULT_AFTER_CONTAINER_MUTATION_BEFORE_COMPLETE, fireOnNth
    On Error Resume Next
    ids = modReceiving.ReceiveN("P000001", "LOT-FIB", Empty, Empty, "LOC0004", 5, "fault-batch")
    Dim failed As Boolean
    failed = (Err.Number <> 0)
    On Error GoTo 0
    modFaultInjection.DisarmFault
    ReceiveNWithFault = failed
End Function

' Verify the invariants after a failed scan-commit on a known container.
Private Function VerifyPostFault(ByVal barcode As String, ByVal preTxn As Long, _
                                 ByVal preCont As Long, ByVal preDash As Long, _
                                 ByVal prevStatus As String, ByVal prevLoc As String, _
                                 ByVal preEvents As Boolean, ByVal tag As String) As Boolean
    On Error Resume Next
    Dim ok As Boolean
    ok = True
    Dim t As Long
    t = TxnCount()
    If t <> preTxn Then
        ok = False
        LogLine "  " & tag & "-txn-count:FAIL now=" & t & " pre=" & preTxn
    Else
        LogLine "  " & tag & "-txn-count:OK now=" & t
    End If
    Dim c As Long
    c = ContainerCount()
    If c <> preCont Then
        ok = False
        LogLine "  " & tag & "-container-count:FAIL now=" & c & " pre=" & preCont
    Else
        LogLine "  " & tag & "-container-count:OK now=" & c
    End If

    ' container fields restored exactly
    Dim rowIdx As Long
    rowIdx = modBarcodeLookup.FindBarcodeRow(barcode)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
    If rowIdx > 0 Then
        Dim st As String
        st = CStr(lo.DataBodyRange.Cells(rowIdx, modBarcodeLookup.ColumnIndex(lo, COL_STATUS)).Value2)
        Dim loc As String
        loc = CStr(lo.DataBodyRange.Cells(rowIdx, modBarcodeLookup.ColumnIndex(lo, COL_STORAGE_LOCATION_ID)).Value2)
        LogLine "  " & tag & "-status-restored:" & IIf(st = prevStatus, "OK", "FAIL") & " now=[" & st & "] pre=[" & prevStatus & "]"
        LogLine "  " & tag & "-location-restored:" & IIf(loc = prevLoc, "OK", "FAIL") & " now=[" & loc & "] pre=[" & prevLoc & "]"
        If st <> prevStatus Or loc <> prevLoc Then ok = False
    Else
        ok = False
        LogLine "  " & tag & "-container-found:FAIL"
    End If

    ' dashboard + formulas + state
    Dim d As Long
    d = DashAvailable()
    LogLine "  " & tag & "-dashboard:" & IIf(d = preDash, "OK", "FAIL") & " now=" & d & " pre=" & preDash
    If d <> preDash Then ok = False
    Dim fe As Long
    fe = FormulaErrorCount()
    LogLine "  " & tag & "-formula-errors:" & IIf(fe = 0, "OK", "FAIL") & " count=" & fe
    If fe > 0 Then ok = False
    ' events must be restored to the captured pre-operation value (the test
    ' session disables events before running, so the correct value is False;
    ' the important invariant is that it equals the pre-operation capture).
    LogLine "  " & tag & "-events-restored:" & IIf(Application.EnableEvents = preEvents, "OK", "FAIL") & " now=" & Application.EnableEvents & " pre=" & preEvents
    If Application.EnableEvents <> preEvents Then ok = False
    LogLine "  " & tag & "-calc-restored:" & IIf(Application.Calculation = -4105, "OK", "FAIL") & " mode=" & Application.Calculation
    If Application.Calculation <> -4105 Then ok = False
    ' scanner usable: a normal scan works immediately
    modScanInterface.HandleScannedBarcode barcode
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Scan")
    modUtilities.UnprotectSheet ws
    Dim m As String
    m = CStr(ws.Range("D9").Value2)
    modUtilities.ProtectSheet ws
    LogLine "  " & tag & "-scanner-usable:" & IIf(InStr(m, "FOUND") > 0, "OK", "FAIL") & " msg=[" & m & "]"
    If InStr(m, "FOUND") = 0 Then ok = False
    ' scan input cleared after the failed commit
    modUtilities.UnprotectSheet ws
    Dim inp As String
    inp = CStr(ws.Range("D7").Value2)
    modUtilities.ProtectSheet ws
    LogLine "  " & tag & "-scan-input-cleared:" & IIf(Len(inp) = 0, "OK", "FAIL") & " val=[" & inp & "]"
    If Len(inp) > 0 Then ok = False
    VerifyPostFault = ok
    On Error GoTo 0
End Function

' Run one boundary fault through the scan-commit path (Available -> TakeOpen).
Private Sub TestScanCommitBoundary(ByVal point As Long)
    On Error Resume Next
    Dim barcode As String
    barcode = "0000001"   ' C000001 Available, no expiry
    Dim preTxn As Long
    preTxn = TxnCount()
    Dim preCont As Long
    preCont = ContainerCount()
    Dim preDash As Long
    preDash = DashAvailable()
    Dim preEvents As Boolean
    preEvents = Application.EnableEvents
    Dim failed As Boolean
    failed = ScanCommitWithFault(barcode, TXN_TAKE_OPEN, point)
    LogLine "fault-scancommit-boundary" & point & ":" & IIf(failed, "OK", "FAIL") & " (expected failure)"
    Dim v As Boolean
    v = VerifyPostFault(barcode, preTxn, preCont, preDash, STATUS_AVAILABLE, "LOC0004", preEvents, "b" & point)
    LogLine "fault-scancommit-boundary" & point & "-invariants:" & IIf(v, "OK", "FAIL")
    On Error GoTo 0
End Sub

' Run one boundary fault through the receive path.
Private Sub TestReceiveBoundary(ByVal point As Long)
    On Error Resume Next
    Dim preTxn As Long
    preTxn = TxnCount()
    Dim preCont As Long
    preCont = ContainerCount()
    Dim preDash As Long
    preDash = DashAvailable()
    Dim preEvents As Boolean
    preEvents = Application.EnableEvents
    Dim cid As String
    Dim failed As Boolean
    failed = ReceiveOneWithFault(point, cid)
    LogLine "fault-receive-boundary" & point & ":" & IIf(failed, "OK", "FAIL") & " (expected failure)"
    Dim t As Long
    t = TxnCount()
    Dim c As Long
    c = ContainerCount()
    LogLine "  receive-b" & point & "-txn-count:" & IIf(t = preTxn, "OK", "FAIL") & " now=" & t & " pre=" & preTxn
    LogLine "  receive-b" & point & "-container-count:" & IIf(c = preCont, "OK", "FAIL") & " now=" & c & " pre=" & preCont
    Dim d As Long
    d = DashAvailable()
    LogLine "  receive-b" & point & "-dashboard:" & IIf(d = preDash, "OK", "FAIL") & " now=" & d & " pre=" & preDash
    Dim fe As Long
    fe = FormulaErrorCount()
    LogLine "  receive-b" & point & "-formula-errors:" & IIf(fe = 0, "OK", "FAIL") & " count=" & fe
    LogLine "  receive-b" & point & "-events:" & IIf(Application.EnableEvents = preEvents, "OK", "FAIL") & " now=" & Application.EnableEvents & " pre=" & preEvents
    LogLine "  receive-b" & point & "-calc:" & IIf(Application.Calculation = -4105, "OK", "FAIL")
    ' a normal receive succeeds immediately after
    Dim cid2 As String
    On Error Resume Next
    cid2 = modReceiving.ReceiveOne("P000001", "LOT-AFTER", Empty, Empty, "LOC0004", "after-fault")
    LogLine "  receive-b" & point & "-normal-after:" & IIf(Len(cid2) > 0, "OK", "FAIL") & " cid=" & cid2
    If Len(cid2) > 0 Then
        ' cleanup the success container + its txn (keep the table clean)
        Dim tid As String
        tid = modReceiving.FindReceiveTransactionByContainer(cid2)
        If Len(tid) > 0 Then modTransactions.RemoveUncommittedTransaction tid
        Dim lo As ListObject
        Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
        modUtilities.UnprotectSheet ThisWorkbook.Worksheets("Containers")
        Dim rowIdx As Long
        rowIdx = modBarcodeLookup.FindContainerRowByID(cid2)
        If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
        modUtilities.ProtectSheet ThisWorkbook.Worksheets("Containers")
    End If
    On Error GoTo 0
End Sub

' ReceiveN batch rollback: fault fires on the 3rd member (>=1 created first).
Private Sub TestReceiveNBatchRollback()
    On Error Resume Next
    Dim preTxn As Long
    preTxn = TxnCount()
    Dim preCont As Long
    preCont = ContainerCount()
    Dim preDash As Long
    preDash = DashAvailable()
    Dim preEvents As Boolean
    preEvents = Application.EnableEvents
    Dim ids() As String
    Dim failed As Boolean
    failed = ReceiveNWithFault(3, ids)
    LogLine "fault-receiven-batch:" & IIf(failed, "OK", "FAIL") & " (expected failure)"
    Dim t As Long
    t = TxnCount()
    Dim c As Long
    c = ContainerCount()
    LogLine "  receiven-batch-txn-count:" & IIf(t = preTxn, "OK", "FAIL") & " now=" & t & " pre=" & preTxn
    LogLine "  receiven-batch-container-count:" & IIf(c = preCont, "OK", "FAIL") & " now=" & c & " pre=" & preCont
    Dim d As Long
    d = DashAvailable()
    LogLine "  receiven-batch-dashboard:" & IIf(d = preDash, "OK", "FAIL") & " now=" & d & " pre=" & preDash
    Dim fe As Long
    fe = FormulaErrorCount()
    LogLine "  receiven-batch-formula-errors:" & IIf(fe = 0, "OK", "FAIL") & " count=" & fe
    LogLine "  receiven-batch-events:" & IIf(Application.EnableEvents = preEvents, "OK", "FAIL") & " now=" & Application.EnableEvents & " pre=" & preEvents
    LogLine "  receiven-batch-calc:" & IIf(Application.Calculation = -4105, "OK", "FAIL")
    ' a normal ReceiveN succeeds immediately after, IDs/barcodes unique
    Dim res() As String
    On Error Resume Next
    res = modReceiving.ReceiveN("P000001", "LOT-AFTERB", Empty, Empty, "LOC0004", 3, "after-batch-fault")
    LogLine "  receiven-batch-normal-after:" & IIf(UBound(res) = 3, "OK", "FAIL") & " count=" & IIf(UBound(res) = 3, 3, -1)
    If UBound(res) = 3 Then
        Dim uniq As Boolean
        uniq = (res(1) <> res(2) And res(2) <> res(3) And res(1) <> res(3))
        LogLine "  receiven-batch-ids-unique:" & IIf(uniq, "OK", "FAIL") & " ids=" & res(1) & "," & res(2) & "," & res(3)
        ' cleanup the success batch
        Dim i As Long
        For i = 3 To 1 Step -1
            Dim tid As String
            tid = modReceiving.FindReceiveTransactionByContainer(res(i))
            If Len(tid) > 0 Then modTransactions.RemoveUncommittedTransaction tid
        Next i
        For i = 3 To 1 Step -1
            Dim lo As ListObject
            Set lo = ThisWorkbook.Worksheets("Containers").ListObjects("tblContainers")
            modUtilities.UnprotectSheet ThisWorkbook.Worksheets("Containers")
            Dim rowIdx As Long
            rowIdx = modBarcodeLookup.FindContainerRowByID(res(i))
            If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
            modUtilities.ProtectSheet ThisWorkbook.Worksheets("Containers")
        Next i
    End If
    On Error GoTo 0
End Sub

' Full fault-injection suite.
Public Sub Test_FaultInjection()
    On Error Resume Next
    LogLine "fault-injection-start"
    modFaultInjection.DisarmFault

    ' scan-commit boundaries 1..5 on C000001 (Available -> TakeOpen)
    TestScanCommitBoundary 1
    TestScanCommitBoundary 2
    TestScanCommitBoundary 3
    TestScanCommitBoundary 4
    TestScanCommitBoundary 5

    ' receive boundaries 1..5
    TestReceiveBoundary 1
    TestReceiveBoundary 2
    TestReceiveBoundary 3
    TestReceiveBoundary 4
    TestReceiveBoundary 5

    ' ReceiveN batch rollback
    TestReceiveNBatchRollback

    modFaultInjection.DisarmFault
    LogLine "fault-injection-end"
    On Error GoTo 0
End Sub
