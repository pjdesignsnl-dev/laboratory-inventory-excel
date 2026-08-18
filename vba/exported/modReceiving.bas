Attribute VB_Name = "modReceiving"
Option Explicit

' ============================================================================
' modReceiving - Receive one / Receive N identical containers (Stages 7-8)
' ============================================================================
' Each receive produces: N unique ContainerIDs, N unique Barcodes, N Receive
' transactions. Atomicity: all-or-nothing via modContainers/modTransactions.
' ============================================================================

Public Function ReceiveOne(ByVal productID As String, ByVal lot As String, _
                           ByVal expiryDate As Variant, ByVal retestDate As Variant, _
                           ByVal locationID As String, _
                           Optional ByVal notes As String = "") As String
    ' Returns the ContainerID of the received container.
    Dim msg As String
    If Not modValidation.ValidateRequiredReceive(productID, lot, locationID, msg) Then
        Err.Raise vbObjectError + 2401, "modReceiving", msg
    End If

    Dim prevEv As Boolean, prevSc As Boolean, prevDa As Boolean
    modUtilities.SaveAppState prevEv, prevSc, prevDa
    Application.EnableEvents = False

    Dim newContainerID As String
    newContainerID = ""

    On Error GoTo rollback
    ' prepare
    Dim barcode As String
    barcode = modContainers.NextBarcode()
    Dim cid As String
    cid = modContainers.NextContainerID()

    ' commit: create container row, then append Receive transaction
    newContainerID = modContainers.AddContainer(barcode, productID, lot, expiryDate, retestDate, locationID, notes)

    Dim snap As Variant
    snap = modTransactions.BuildSnapshot(TXN_RECEIVE, barcode, cid, productID, _
                                         modBarcodeLookup.productName(productID), _
                                         PREV_NONE, STATUS_AVAILABLE, PREV_NONE, locationID, _
                                         lot, "", "", notes)
    Call modTransactions.AppendTransaction(snap)

    ' post-condition: container exists exactly once, txn exists exactly once
    If Not modContainers.ContainerIDExists(cid) Then Err.Raise vbObjectError + 2402, "modReceiving", "Container post-condition failed"
    If Not modTransactions.VerifyAppended(CStr(snap(1))) Then Err.Raise vbObjectError + 2403, "modReceiving", "Transaction post-condition failed"

    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    ReceiveOne = cid
    Exit Function

rollback:
    ' remove the container row we just created and any transaction row
    On Error Resume Next
    If Len(newContainerID) > 0 Then
        Dim lo As ListObject
        Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
        modUtilities.UnprotectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
        Dim rowIdx As Long
        rowIdx = modBarcodeLookup.FindContainerRowByID(newContainerID)
        If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
        modUtilities.ProtectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
    End If
    On Error GoTo 0
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Err.Raise vbObjectError + 2404, "modReceiving", "Receive failed and was rolled back: " & Err.Description
End Function

Public Function ReceiveN(ByVal productID As String, ByVal lot As String, _
                         ByVal expiryDate As Variant, ByVal retestDate As Variant, _
                         ByVal locationID As String, ByVal count As Long, _
                         Optional ByVal notes As String = "") As String()
    ' Receives 'count' identical containers. Returns an array of ContainerIDs.
    ' All-or-nothing: on any failure, all created containers + transactions are
    ' rolled back and the error is raised.
    Dim results() As String
    ReDim results(1 To count)

    Dim msg As String
    If Not modValidation.ValidateRequiredReceive(productID, lot, locationID, msg) Then
        Err.Raise vbObjectError + 2401, "modReceiving", msg
    End If
    If count < 1 Or count > 999 Then
        Err.Raise vbObjectError + 2405, "modReceiving", "Quantity must be between 1 and 999."
    End If

    Dim prevEv As Boolean, prevSc As Boolean, prevDa As Boolean
    modUtilities.SaveAppState prevEv, prevSc, prevDa
    Application.EnableEvents = False

    Dim created As Collection
    Set created = New Collection

    On Error GoTo rollbackN
    Dim i As Long
    For i = 1 To count
        Dim cid As String
        cid = ReceiveOne(productID, lot, expiryDate, retestDate, locationID, notes)
        created.Add cid
        results(i) = cid
    Next i

    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    ReceiveN = results
    Exit Function

rollbackN:
    On Error Resume Next
    Dim j As Long
    For j = created.count To 1 Step -1
        Dim lo As ListObject
        Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
        modUtilities.UnprotectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
        Dim rowIdx As Long
        rowIdx = modBarcodeLookup.FindContainerRowByID(CStr(created(j)))
        If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
        modUtilities.ProtectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
    Next j
    On Error GoTo 0
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Err.Raise vbObjectError + 2406, "modReceiving", "Batch receive failed and was rolled back: " & Err.Description
End Function
