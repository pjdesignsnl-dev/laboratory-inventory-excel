Attribute VB_Name = "modScanInterface"
Option Explicit

' ============================================================================
' modScanInterface - Scan input handling, transaction commit, reset/focus
' ============================================================================
' Dispatch point for the Scan worksheet event. Implements the duplicate-scan
' guard (re-entrancy protection for the same pending operation) and the
' scan -> lookup -> validate -> commit -> reset -> focus pipeline.
' ============================================================================

Private m_busy As Boolean
Private m_lastOperationToken As String

Public Property Get IsBusy() As Boolean
    IsBusy = m_busy
End Property

' ------------------------------------------------------------------ entry point (called by Scan.Worksheet_Change)
Public Sub HandleScannedBarcode(ByVal rawValue As Variant)
    ' Re-entrancy / duplicate-scan guard: if a scan operation is already in
    ' flight, ignore. This prevents repeated Enter events and event re-entry
    ' from double-processing the SAME pending scan. Legitimate later scans of
    ' the same container are allowed once the operation completes/aborts.
    If m_busy Then Exit Sub
    m_busy = True

    Dim prevEv As Boolean, prevSc As Boolean, prevDa As Boolean
    modUtilities.SaveAppState prevEv, prevSc, prevDa
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo failOut

    Dim barcode As String
    barcode = modUtilities.NormalizeBarcode(rawValue)

    Dim rowNum As Long
    rowNum = modBarcodeLookup.FindBarcodeRow(barcode)
    Dim result As Long
    result = modBarcodeLookup.LookupState(barcode)

    ' populate staging (frozen tblScanResults)
    modBarcodeLookup.PopulateScanStaging barcode, rowNum, result

    ' status message
    Dim msg As String
    Select Case result
        Case modBarcodeLookup.LR_EMPTY
            msg = "Scan a barcode."
        Case modBarcodeLookup.LR_INVALID
            msg = "Invalid barcode format (expected 7 digits)."
        Case modBarcodeLookup.LR_UNKNOWN
            msg = "UNKNOWN BARCODE - receive this container first."
        Case modBarcodeLookup.LR_DUPLICATE
            msg = "DUPLICATE BARCODE - multiple containers share this barcode."
        Case modBarcodeLookup.LR_FOUND
            msg = "FOUND - scan details shown. Choose an action to commit."
    End Select
    SetStatusMessage msg

    m_busy = False
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Exit Sub

failOut:
    SetStatusMessage "Scan error: " & Err.Description
    m_busy = False
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
End Sub

' ------------------------------------------------------------------ pending operation commit
Public Sub CommitAction(ByVal txnType As String, Optional ByVal newLocation As String = "", _
                        Optional ByVal reason As String = "", Optional ByVal notes As String = "")
    ' Executes a mutation for the currently scanned container using the frozen
    ' transition matrix. Uses atomic commit via modContainers + modTransactions.
    If m_busy Then Exit Sub
    m_busy = True

    Dim prevEv As Boolean, prevSc As Boolean, prevDa As Boolean
    modUtilities.SaveAppState prevEv, prevSc, prevDa
    Application.EnableEvents = False

    On Error GoTo failCommit

    Dim scanWs As Worksheet
    Set scanWs = ThisWorkbook.Worksheets(WS_SCAN)
    Dim staging As ListObject
    Set staging = scanWs.ListObjects(TBL_SCAN_RESULTS)

    Dim barcode As String
    barcode = modUtilities.NormalizeBarcode(staging.DataBodyRange.Cells(1, modBarcodeLookup.ColumnIndex(staging, "Barcode")).Value2)
    If Len(barcode) = 0 Then
        SetStatusMessage "No container scanned. Scan a barcode first."
        m_busy = False
        modUtilities.RestoreAppState prevEv, prevSc, prevDa
        Exit Sub
    End If

    Dim rowNum As Long
    rowNum = modBarcodeLookup.FindBarcodeRow(barcode)
    Dim result As Long
    result = modBarcodeLookup.LookupState(barcode)
    If result <> modBarcodeLookup.LR_FOUND Then
        SetStatusMessage "Cannot commit: barcode not uniquely found."
        m_busy = False
        modUtilities.RestoreAppState prevEv, prevSc, prevDa
        Exit Sub
    End If

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim status As String
    status = CStr(lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_STATUS)).Value2)
    Dim expiry As Variant
    expiry = lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_EXPIRY_DATE)).Value2

    ' validate transition
    Dim vmsg As String
    Dim vclass As MsgClass
    If Not modValidation.ValidateTransition(status, txnType, expiry, vmsg, vclass) Then
        SetStatusMessage vmsg
        m_busy = False
        modUtilities.RestoreAppState prevEv, prevSc, prevDa
        Exit Sub
    End If
    If vclass = mcConfirm Then
        ' In the macro-free UI, confirmation is the operator's explicit choice
        ' to invoke this action; treat as confirmed.
    End If

    ' determine new state
    Dim newStatus As String
    Dim newLoc As String
    Dim setOpened As Boolean
    Dim setDisposed As Boolean
    Dim dReason As String
    DetermineEffects txnType, status, newLocation, newStatus, newLoc, setOpened, setDisposed, dReason

    ' capture old state (atomicity)
    Dim saved As modContainers.ContainerState
    saved = modContainers.CaptureState(rowNum)

    ' PREPARE: build the snapshot BEFORE any mutation (atomic boundary)
    Dim snap As Variant
    snap = modTransactions.BuildSnapshot(txnType, barcode, _
        CStr(lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_CONTAINER_ID)).Value2), _
        CStr(lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_PRODUCT_ID)).Value2), _
        modBarcodeLookup.ProductName(CStr(lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_PRODUCT_ID)).Value2)), _
        saved.Status, newStatus, _
        saved.StorageLocationID, newLoc, _
        CStr(lo.DataBodyRange.Cells(rowNum, modBarcodeLookup.ColumnIndex(lo, COL_BATCH_LOT)).Value2), _
        reason, "", notes)

    On Error GoTo rollbackCommit

    ' COMMIT: append transaction first, then mutate container
    Dim tid As String
    tid = modTransactions.AppendTransaction(snap)
    modContainers.ApplyStateChange rowNum, newStatus, newLoc, setOpened, setDisposed, dReason

    ' post-conditions
    If Not modTransactions.VerifyAppended(tid) Then Err.Raise vbObjectError + 2501, "modScanInterface", "Transaction post-condition failed"
    If Not modContainers.BarcodeExists(barcode) Then Err.Raise vbObjectError + 2502, "modScanInterface", "Container post-condition failed"

    SetStatusMessage "OK: " & txnType & " committed (" & tid & ")."
    m_busy = False
    ResetScanField
    FocusScanInput
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Exit Sub

rollbackCommit:
    ' atomic rollback: restore container fields and remove the uncommitted
    ' transaction row (only if it belongs to this operation and was not
    ' accepted externally).
    On Error Resume Next
    modContainers.RollbackState saved
    If Len(tid) > 0 Then modTransactions.RemoveUncommittedTransaction tid
    On Error GoTo 0
    SetStatusMessage "Commit failed and was rolled back: " & Err.Description
    m_busy = False
    ResetScanField
    FocusScanInput
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Exit Sub

failCommit:
    SetStatusMessage "Action error: " & Err.Description
    m_busy = False
    ResetScanField
    FocusScanInput
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
End Sub

' ------------------------------------------------------------------ effects
Private Sub DetermineEffects(ByVal txnType As String, ByVal currentStatus As String, _
                             ByVal requestedLocation As String, _
                             ByRef newStatus As String, ByRef newLoc As String, _
                             ByRef setOpened As Boolean, ByRef setDisposed As Boolean, _
                             ByRef dReason As String)
    newStatus = currentStatus
    newLoc = requestedLocation
    setOpened = False
    setDisposed = False
    dReason = ""

    Select Case txnType
        Case TXN_TAKE_OPEN
            newStatus = STATUS_IN_USE: newLoc = "": setOpened = True
        Case TXN_RETURN
            newStatus = STATUS_AVAILABLE
        Case TXN_TRANSFER
            ' status unchanged; newLoc already set
        Case TXN_DISPOSE
            newStatus = STATUS_DISPOSED: newLoc = "": setDisposed = True: dReason = modUtilities.Coalesce(dReason, REASON_OTHER)
        Case TXN_MARK_EXPIRED
            newStatus = STATUS_EXPIRED
        Case TXN_MARK_DAMAGED
            newStatus = STATUS_DAMAGED
        Case TXN_MARK_MISSING
            newStatus = STATUS_MISSING: newLoc = ""
        Case TXN_ADJUSTMENT
            ' caller sets newStatus/newLoc explicitly via requestedLocation convention
    End Select
End Sub

' ------------------------------------------------------------------ UI helpers
Public Sub SetStatusMessage(ByVal msg As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_SCAN)
    modUtilities.UnprotectSheet ws
    ws.Range("D9").Value2 = msg
    modUtilities.ProtectSheet ws
    On Error GoTo 0
End Sub

Public Sub ResetScanField()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_SCAN)
    modUtilities.UnprotectSheet ws
    ws.Range("D7").Value2 = ""
    modUtilities.ProtectSheet ws
    On Error GoTo 0
End Sub

Public Sub FocusScanInput()
    ' Narrowly documented UI-focus operation (Phase B principle 4 exception).
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_SCAN)
    ws.Activate
    ws.Range("D7").Select
    On Error GoTo 0
End Sub
