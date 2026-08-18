Attribute VB_Name = "modValidation"
Option Explicit

' ============================================================================
' modValidation - Transition rules, D-018 expiry semantics, required fields
' ============================================================================
' Implements the frozen state-transition matrix (docs/status-transition-matrix.md)
' and the D-018 rule: an Available container that is expired by date is
' excluded from usable stock and blocked from TakeOpen; Status is never
' silently mutated by the clock.
' ============================================================================

Public Enum TransitionDecision
    tdAllow = 0
    tdConfirm = 1
    tdBlock = 2
    tdNotApplicable = 3
End Enum

' ------------------------------------------------------------------ transition matrix
Public Function ValidateTransition(ByVal currentStatus As String, _
                                   ByVal txnType As String, _
                                   ByVal expiryDate As Variant, _
                                   ByRef message As String, _
                                   ByRef msgClass As MsgClass) As Boolean
    ' Returns True if the transition is permitted (ALLOW or CONFIRM).
    ' expiryDate is the container's ExpiryDate (may be empty).
    Dim d As TransitionDecision
    d = DecisionFor(currentStatus, txnType, expiryDate)

    Select Case d
        Case tdAllow
            ValidateTransition = True
            msgClass = mcInfo
            message = ""
        Case tdConfirm
            ValidateTransition = True
            msgClass = mcConfirm
            message = "This action requires confirmation."
        Case tdBlock
            ValidateTransition = False
            msgClass = mcBlocking
            message = BlockMessage(currentStatus, txnType)
        Case Else
            ValidateTransition = False
            msgClass = mcBlocking
            message = "Transaction '" & txnType & "' does not apply to status '" & currentStatus & "'."
    End Select
End Function

Private Function IsExpiredByDate(ByVal expiryDate As Variant) As Boolean
    If IsBlankOrEmpty(expiryDate) Then Exit Function
    If Not IsDate(expiryDate) Then Exit Function
    If CDate(expiryDate) < modUtilities.GetNow() Then IsExpiredByDate = True
End Function

Private Function DecisionFor(ByVal status As String, ByVal txn As String, ByVal expiryDate As Variant) As TransitionDecision
    Dim exp As Boolean
    exp = IsExpiredByDate(expiryDate)

    ' ---------- Receive: only on a new (non-existent) container
    If status = PREV_NONE Then
        If txn = TXN_RECEIVE Then DecisionFor = tdAllow Else DecisionFor = tdNotApplicable
        Exit Function
    End If

    Select Case status
        Case STATUS_AVAILABLE
            If exp Then
                ' D-018: expired-by-date Available cannot TakeOpen; may Dispose,
                ' Transfer (relocation), or MarkExpired.
                Select Case txn
                    Case TXN_TAKE_OPEN: DecisionFor = tdBlock
                    Case TXN_TRANSFER: DecisionFor = tdAllow
                    Case TXN_DISPOSE: DecisionFor = tdAllow
                    Case TXN_MARK_EXPIRED: DecisionFor = tdAllow
                    Case TXN_MARK_DAMAGED: DecisionFor = tdConfirm
                    Case TXN_MARK_MISSING: DecisionFor = tdConfirm
                    Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                    Case Else: DecisionFor = tdBlock
                End Select
            Else
                Select Case txn
                    Case TXN_TAKE_OPEN: DecisionFor = tdAllow
                    Case TXN_TRANSFER: DecisionFor = tdAllow
                    Case TXN_DISPOSE: DecisionFor = tdConfirm
                    Case TXN_MARK_EXPIRED: DecisionFor = tdConfirm
                    Case TXN_MARK_DAMAGED: DecisionFor = tdConfirm
                    Case TXN_MARK_MISSING: DecisionFor = tdConfirm
                    Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                    Case Else: DecisionFor = tdBlock
                End Select
            End If

        Case STATUS_IN_USE
            Select Case txn
                Case TXN_RETURN: DecisionFor = tdAllow
                Case TXN_DISPOSE: DecisionFor = tdConfirm
                Case TXN_MARK_EXPIRED: DecisionFor = tdConfirm
                Case TXN_MARK_DAMAGED: DecisionFor = tdConfirm
                Case TXN_MARK_MISSING: DecisionFor = tdConfirm
                Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                Case Else: DecisionFor = tdBlock
            End Select

        Case STATUS_EXPIRED
            Select Case txn
                Case TXN_DISPOSE: DecisionFor = tdAllow
                Case TXN_TRANSFER: DecisionFor = tdAllow
                Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                Case Else: DecisionFor = tdBlock
            End Select

        Case STATUS_DAMAGED
            Select Case txn
                Case TXN_DISPOSE: DecisionFor = tdAllow
                Case TXN_TRANSFER: DecisionFor = tdAllow
                Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                Case Else: DecisionFor = tdBlock
            End Select

        Case STATUS_DISPOSED
            If txn = TXN_ADJUSTMENT Then
                DecisionFor = tdConfirm
            Else
                DecisionFor = tdBlock
            End If

        Case STATUS_MISSING
            Select Case txn
                Case TXN_DISPOSE: DecisionFor = tdConfirm
                Case TXN_ADJUSTMENT: DecisionFor = tdConfirm
                Case Else: DecisionFor = tdBlock
            End Select

        Case Else
            DecisionFor = tdBlock
    End Select
End Function

Private Function BlockMessage(ByVal status As String, ByVal txn As String) As String
    If status = STATUS_AVAILABLE And txn = TXN_TAKE_OPEN Then
        BlockMessage = "TakeOpen blocked: container is expired by date (D-018). Record MarkExpired or Dispose."
        Exit Function
    End If
    If status = STATUS_DISPOSED Then
        BlockMessage = "Disposed containers cannot be reactivated through normal transactions."
        Exit Function
    End If
    If status = STATUS_MISSING And txn = TXN_TAKE_OPEN Then
        BlockMessage = "Missing containers cannot be taken through a normal stock transaction."
        Exit Function
    End If
    If status = STATUS_EXPIRED And txn = TXN_TAKE_OPEN Then
        BlockMessage = "Expired containers cannot be taken through a normal stock transaction."
        Exit Function
    End If
    BlockMessage = "Transition '" & txn & "' is not allowed for status '" & status & "'."
End Function

' ------------------------------------------------------------------ required-field validation
Public Function ValidateRequiredReceive(ByVal productID As String, ByVal lot As String, _
                                        ByVal locationID As String, ByRef message As String) As Boolean
    If modUtilities.IsBlankOrEmpty(productID) Then
        message = "Product ID is required."
        Exit Function
    End If
    If modUtilities.IsBlankOrEmpty(lot) Then
        message = "Batch/Lot number is required."
        Exit Function
    End If
    If modUtilities.IsBlankOrEmpty(locationID) Then
        message = "Storage location is required."
        Exit Function
    End If
    If Not ProductExists(productID) Then
        message = "Product '" & productID & "' does not exist."
        Exit Function
    End If
    If Not LocationExists(locationID) Then
        message = "Location '" & locationID & "' does not exist."
        Exit Function
    End If
    ValidateRequiredReceive = True
End Function

Public Function ProductExists(ByVal productID As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_PRODUCTS).ListObjects(TBL_PRODUCTS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_PRODUCT_ID)
    If idCol <= 0 Then Exit Function
    On Error Resume Next
    ProductExists = Not IsError(Application.Match(productID, lo.DataBodyRange.Columns(idCol), 0))
    On Error GoTo 0
End Function

Public Function LocationExists(ByVal locationID As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_LOCATIONS).ListObjects(TBL_LOCATIONS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_STORAGE_LOCATION_ID)
    If idCol <= 0 Then Exit Function
    On Error Resume Next
    LocationExists = Not IsError(Application.Match(locationID, lo.DataBodyRange.Columns(idCol), 0))
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ D-018 stock helper
Public Function IsUsableAvailable(ByVal status As String, ByVal expiryDate As Variant) As Boolean
    ' Usable available stock = Status="Available" AND (no expiry OR expiry >= today).
    If status <> STATUS_AVAILABLE Then Exit Function
    If IsBlankOrEmpty(expiryDate) Then
        IsUsableAvailable = True
        Exit Function
    End If
    If IsDate(expiryDate) And CDate(expiryDate) >= modUtilities.GetNow() Then IsUsableAvailable = True
End Function
