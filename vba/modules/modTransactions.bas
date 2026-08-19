Attribute VB_Name = "modTransactions"
Option Explicit

' ============================================================================
' modTransactions - Transaction ID generation, snapshot builder, atomic append
' ============================================================================
' Implements stages 4 and 5 of the plan:
'  - NextTransactionID(): MAX over the frozen TransactionID helper semantics
'    (T########, uniqueness verified before commit).
'  - BuildSnapshot(): immutable snapshot of the operation.
'  - AppendTransaction(): atomic append with post-condition verification.
' ============================================================================

' ------------------------------------------------------------------ ID generation
Public Function NextTransactionID() As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If idCol <= 0 Then Err.Raise vbObjectError + 2201, "modTransactions", "TransactionID column not found"

    Dim maxN As Long
    maxN = 0
    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.Count
            Dim v As Variant
            v = lo.DataBodyRange.Cells(r, idCol).Value2
            If Not IsEmpty(v) And IsNumeric(Mid$(CStr(v), 2)) Then
                Dim n As Long
                n = CLng(Mid$(CStr(v), 2))
                If n > maxN Then maxN = n
            End If
        Next r
    End If

    ' generate next unique ID (verify not already used)
    Dim candidate As String
    Do
        maxN = maxN + 1
        candidate = modContainers.PadID(maxN, FMT_TRANSACTION_ID)
    Loop While TransactionIDExists(candidate)

    NextTransactionID = candidate
End Function

Public Function TransactionIDExists(ByVal tid As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If idCol <= 0 Then Exit Function
    On Error Resume Next
    TransactionIDExists = Not IsError(Application.Match(tid, lo.DataBodyRange.Columns(idCol), 0))
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ snapshot builder
' BuildSnapshot returns a 1-based array aligned to the frozen column order of
' tblTransactions.
Public Function BuildSnapshot(ByVal txnType As String, _
                              ByVal barcode As String, ByVal containerID As String, _
                              ByVal productID As String, ByVal productName As String, _
                              ByVal previousStatus As String, ByVal newStatus As String, _
                              ByVal previousLocation As String, ByVal newLocation As String, _
                              ByVal lot As String, _
                              Optional ByVal reason As String = "", _
                              Optional ByVal reference As String = "", _
                              Optional ByVal notes As String = "") As Variant
    Dim snap(1 To 16) As Variant
    snap(1) = NextTransactionID()
    snap(2) = modUtilities.GetNow()
    snap(3) = modUtilities.GetOperator()
    snap(4) = barcode
    snap(5) = containerID
    snap(6) = productID
    snap(7) = productName
    snap(8) = txnType
    snap(9) = previousStatus
    snap(10) = newStatus
    snap(11) = previousLocation
    snap(12) = newLocation
    snap(13) = lot
    snap(14) = reason
    snap(15) = reference
    snap(16) = notes
    BuildSnapshot = snap
End Function

' ------------------------------------------------------------------ atomic append
Public Function AppendTransaction(ByRef snap As Variant) As String
    ' Appends the snapshot as a new row in tblTransactions.
    ' Returns the TransactionID. Raises on failure (caller rolls back).
    '
    ' Fault-injection hooks (test-only):
    '   FAULT_BEFORE_TXN_APPEND  -> raised before any row allocation.
    '   FAULT_AFTER_TXN_APPEND   -> raised after the row is fully written but
    '                               before the post-condition; this procedure
    '                               then removes ITS OWN just-created row so
    '                               no orphan can survive even if the caller
    '                               has no tid to remove (self-cleaning).
    On Error GoTo appendFail

    ' BOUNDARY 1: before transaction append (no row allocated yet)
    Call modFaultInjection.FaultAt(modFaultInjection.FAULT_BEFORE_TXN_APPEND)

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_TRANSACTIONS)
    modUtilities.UnprotectSheet ws

    Dim lr As ListRow
    Set lr = lo.ListRows.Add

    Dim colCount As Long
    colCount = lo.ListColumns.Count
    Dim i As Long
    For i = 1 To colCount
        lr.Range.Cells(1, i).Value2 = snap(i)
    Next i

    Dim tid As String
    tid = CStr(snap(1))

    ' BOUNDARY 2: immediately after the transaction row is allocated+written.
    ' If the fault fires, delete the row we just created (self-clean), then
    ' raise so the caller still observes failure.
    If modFaultInjection.FaultAt(modFaultInjection.FAULT_AFTER_TXN_APPEND) Then
        On Error Resume Next
        RemoveRowById tid
        On Error GoTo appendFail
        modUtilities.ProtectSheet ws
        Err.Raise vbObjectError + 2202, "modTransactions", _
                  "INJECTED FAULT after transaction append (row removed)"
    End If

    modUtilities.ProtectSheet ws

    ' post-condition: exactly one row with this TransactionID
    If Not VerifyAppended(tid) Then
        Err.Raise vbObjectError + 2202, "modTransactions", _
                  "Post-condition failed: transaction '" & tid & "' not appended exactly once."
    End If

    AppendTransaction = tid
    Exit Function

appendFail:
    ' If we raised after allocating a row (boundary 2 or post-condition),
    ' remove the row that belongs to THIS uncommitted operation only.
    On Error Resume Next
    If Len(tid) > 0 Then
        RemoveRowById tid
        modUtilities.ProtectSheet ws
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 2203, "modTransactions", _
              "Transaction append failed and was rolled back: " & Err.Description
End Function

Private Sub RemoveRowById(ByVal tid As String)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If idCol <= 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_TRANSACTIONS)
    modUtilities.UnprotectSheet ws
    Dim m As Variant
    On Error Resume Next
    m = Application.Match(tid, lo.DataBodyRange.Columns(idCol), 0)
    On Error GoTo 0
    If Not IsError(m) Then lo.ListRows(CLng(m)).Delete
    modUtilities.ProtectSheet ws
End Sub

Public Function VerifyAppended(ByVal tid As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If idCol <= 0 Then Exit Function
    On Error Resume Next
    VerifyAppended = (Application.CountIf(lo.DataBodyRange.Columns(idCol), tid) = 1)
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ removal of uncommitted transaction (rollback helper)
Public Sub RemoveUncommittedTransaction(ByVal tid As String)
    ' Used ONLY by the atomic rollback: removes the transaction row created by
    ' the current uncommitted operation. Historical committed transactions are
    ' never deleted.
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_TRANSACTIONS).ListObjects(TBL_TRANSACTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_TRANSACTION_ID)
    If idCol <= 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_TRANSACTIONS)
    modUtilities.UnprotectSheet ws

    Dim m As Variant
    On Error Resume Next
    m = Application.Match(tid, lo.DataBodyRange.Columns(idCol), 0)
    On Error GoTo 0
    If Not IsError(m) Then
        lo.ListRows(CLng(m)).Delete
    End If

    modUtilities.ProtectSheet ws
End Sub
