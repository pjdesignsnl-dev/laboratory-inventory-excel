Attribute VB_Name = "modReceiving"
Option Explicit

' ============================================================================
' modReceiving - Receive one / Receive N identical containers (Stages 7-8)
' ============================================================================
' Each receive produces: N unique ContainerIDs, N unique Barcodes, N Receive
' transactions. Atomicity: all-or-nothing via modContainers/modTransactions.
'
' Fault-injection boundaries (test-only):
'   1 = before transaction append            (raised by AppendTransaction)
'   2 = after transaction append             (raised by AppendTransaction,
'                                             self-cleaning)
'   3 = before container mutation            (raised by ApplyStateChange when
'                                             used via CommitAction; for
'                                             receive the container is created
'                                             by AddContainer)
'   4 = during/after container mutation      (AddContainer self-cleaning)
'   5 = after container mutation, before workflow completion (raised here)
'
' Rollback removes ONLY rows belonging to the current uncommitted operation:
' the container row(s) created here and their Receive transaction row(s).
' Historical committed audit rows are never touched.
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
    Dim newTxnID As String
    newTxnID = ""

    On Error GoTo rollback
    ' prepare
    Dim barcode As String
    barcode = modContainers.NextBarcode()
    Dim cid As String
    cid = modContainers.NextContainerID()

    ' BOUNDARY 3 (receive variant): before container creation/mutation.
    ' Nothing has been appended or created yet; rollback is a no-op.
    Call modFaultInjection.FaultAt(modFaultInjection.FAULT_BEFORE_CONTAINER_MUTATION)

    ' commit: create container row, then append Receive transaction
    newContainerID = modContainers.AddContainer(barcode, productID, lot, expiryDate, retestDate, locationID, notes)

    Dim snap As Variant
    snap = modTransactions.BuildSnapshot(TXN_RECEIVE, barcode, cid, productID, _
                                         modBarcodeLookup.ProductName(productID), _
                                         PREV_NONE, STATUS_AVAILABLE, PREV_NONE, locationID, _
                                         lot, "", "", notes)
    newTxnID = modTransactions.AppendTransaction(snap)

    ' BOUNDARY 5: after container mutation + transaction append, before the
    ' workflow reports success. Caller rollback must restore both.
    Call modFaultInjection.FaultAt(modFaultInjection.FAULT_AFTER_CONTAINER_MUTATION_BEFORE_COMPLETE)

    ' post-condition: container exists exactly once, txn exists exactly once
    If Not modContainers.ContainerIDExists(cid) Then Err.Raise vbObjectError + 2402, "modReceiving", "Container post-condition failed"
    If Not modTransactions.VerifyAppended(newTxnID) Then Err.Raise vbObjectError + 2403, "modReceiving", "Transaction post-condition failed"

    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    ReceiveOne = cid
    Exit Function

rollback:
    ' Atomic rollback of THIS uncommitted receive: remove the container row we
    ' created (if any) and the Receive transaction row we appended (if any).
    ' Historical committed transactions are never deleted.
    On Error Resume Next
    If Len(newTxnID) > 0 Then modTransactions.RemoveUncommittedTransaction newTxnID
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
    ' All-or-nothing: on any failure, all created containers + their Receive
    ' transactions are rolled back and the error is raised.
    On Error GoTo rollbackN

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
    Dim createdTxns As Collection
    Set createdTxns = New Collection

    Dim i As Long
    For i = 1 To count
        Dim cid As String
        cid = ReceiveOne(productID, lot, expiryDate, retestDate, locationID, notes)
        created.Add cid
        results(i) = cid
        ' the Receive transaction for this container (uncommitted batch member)
        Dim tid As String
        tid = modReceiving.FindReceiveTransactionByContainer(cid)
        If Len(tid) > 0 Then createdTxns.Add tid
    Next i

    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    ReceiveN = results
    Exit Function

rollbackN:
    ' All-or-nothing: remove every Receive transaction row appended by this
    ' batch, then every container row created by this batch (reverse order).
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Dim k As Long
    If Not createdTxns Is Nothing Then
        For k = createdTxns.Count To 1 Step -1
            modTransactions.RemoveUncommittedTransaction CStr(createdTxns(k))
        Next k
    End If
    Dim j As Long
    If Not created Is Nothing Then
        For j = created.Count To 1 Step -1
            Dim lo As ListObject
            Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
            modUtilities.UnprotectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
            Dim rowIdx As Long
            rowIdx = modBarcodeLookup.FindContainerRowByID(CStr(created(j)))
            If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
            modUtilities.ProtectSheet ThisWorkbook.Worksheets(WS_CONTAINERS)
        Next j
    End If
    On Error GoTo 0
    modUtilities.RestoreAppState prevEv, prevSc, prevDa
    Err.Raise vbObjectError + 2406, "modReceiving", _
              "Batch receive failed and was rolled back (err " & errNum & " " & errDesc & "): " & Err.Description
End Function

' ------------------------------------------------------------------ helper: receive txn of a container
Public Function FindReceiveTransactionByContainer(ByVal containerID As String) As String
    ' Returns the TransactionID of the most recent Receive transaction for the
    ' given container, or "". Used by ReceiveN rollback to remove the batch's
    ' own Receive rows (never historical rows).
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim cidCol As Long
    cidCol = modBarcodeLookup.ColumnIndex(lo, COL_CONTAINER_ID)
    Dim txnCol As Long
    txnCol = modBarcodeLookup.ColumnIndex(lo, COL_TXN_TYPE)
    Dim tidCol As Long
    tidCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If cidCol <= 0 Or txnCol <= 0 Or tidCol <= 0 Then Exit Function
    ' scan from the bottom (most recent) upward
    Dim r As Long
    For r = lo.DataBodyRange.Rows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(r, cidCol).Value2) = containerID Then
            If CStr(lo.DataBodyRange.Cells(r, txnCol).Value2) = TXN_RECEIVE Then
                FindReceiveTransactionByContainer = CStr(lo.DataBodyRange.Cells(r, tidCol).Value2)
                Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function
