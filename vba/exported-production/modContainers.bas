Attribute VB_Name = "modContainers"
Option Explicit

' ============================================================================
' modContainers - Container state mutation, ID generation, rollback (Stage 6)
' ============================================================================
' Applies a state/location change to a container row, capturing old state for
' rollback. Generates next ContainerID/Barcode using the frozen helper strategy
' (MAX + 1, uniqueness verified).
' ============================================================================

Public Type ContainerState
    Row As Long
    status As String
    StorageLocationID As String
    OpenedDate As Variant
    DisposalDate As Variant
    disposalReason As String
    notes As String
End Type

' ------------------------------------------------------------------ next IDs
Public Function NextContainerID() As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim numCol As Long
    numCol = modBarcodeLookup.ColumnIndex(lo, COL_HELPER_CONTAINER_NUM)
    If numCol <= 0 Then Err.Raise vbObjectError + 2301, "modContainers", "HelperContainerNum column not found"

    Dim maxN As Long
    maxN = 0
    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.count
            Dim v As Variant
            v = lo.DataBodyRange.Cells(r, numCol).Value2
            If IsNumeric(v) And CLng(v) > maxN Then maxN = CLng(v)
        Next r
    End If

    Dim candidate As String
    Do
        maxN = maxN + 1
        candidate = PadID(maxN, FMT_CONTAINER_ID)
    Loop While ContainerIDExists(candidate)
    NextContainerID = candidate
End Function

Public Function PadID(ByVal n As Long, ByVal fmt As String) As String
    ' Manual zero-padded ID: fmt is e.g. "C000000" (prefix C + 6 digits).
    ' Format$(n, fmt) can misparse the digit run as a date in some locales,
    ' so we build the string explicitly.
    Dim prefix As String
    prefix = ""
    Dim digits As Long
    digits = 0
    Dim i As Long
    For i = 1 To Len(fmt)
        If Mid$(fmt, i, 1) >= "0" And Mid$(fmt, i, 1) <= "9" Then
            digits = digits + 1
        Else
            prefix = prefix & Mid$(fmt, i, 1)
        End If
    Next i
    PadID = prefix & Right$(String$(digits, "0") & CStr(n), digits)
End Function

Public Function NextBarcode() As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim numCol As Long
    numCol = modBarcodeLookup.ColumnIndex(lo, COL_HELPER_BARCODE_NUM)
    If numCol <= 0 Then Err.Raise vbObjectError + 2302, "modContainers", "HelperBarcodeNum column not found"

    Dim maxN As Long
    maxN = 0
    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.count
            Dim v As Variant
            v = lo.DataBodyRange.Cells(r, numCol).Value2
            If IsNumeric(v) And CLng(v) > maxN Then maxN = CLng(v)
        Next r
    End If

    Dim candidate As String
    Do
        maxN = maxN + 1
        candidate = PadID(maxN, FMT_BARCODE)
    Loop While BarcodeExists(candidate)
    NextBarcode = candidate
End Function

Public Function ContainerIDExists(ByVal containerID As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long
    idCol = modBarcodeLookup.ColumnIndex(lo, COL_CONTAINER_ID)
    If idCol <= 0 Then Exit Function
    On Error Resume Next
    ContainerIDExists = Not IsError(Application.Match(containerID, lo.DataBodyRange.Columns(idCol), 0))
    On Error GoTo 0
End Function

Public Function BarcodeExists(ByVal barcode As String) As Boolean
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim bcCol As Long
    bcCol = modBarcodeLookup.ColumnIndex(lo, COL_BARCODE)
    If bcCol <= 0 Then Exit Function
    On Error Resume Next
    BarcodeExists = Not IsError(Application.Match(barcode, lo.DataBodyRange.Columns(bcCol), 0))
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ state mutation
Public Function CaptureState(ByVal rowNum As Long) As ContainerState
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim r As Long
    r = lo.DataBodyRange.Row + rowNum - 1
    With CaptureState
        .Row = rowNum
        .status = CellValue(lo, rowNum, COL_STATUS)
        .StorageLocationID = CellValue(lo, rowNum, COL_STORAGE_LOCATION_ID)
        .OpenedDate = CellValue(lo, rowNum, COL_OPENED_DATE)
        .DisposalDate = CellValue(lo, rowNum, COL_DISPOSAL_DATE)
        .disposalReason = CellValue(lo, rowNum, COL_DISPOSAL_REASON)
        .notes = CellValue(lo, rowNum, COL_NOTES)
    End With
End Function

Public Sub ApplyStateChange(ByVal rowNum As Long, _
                            ByVal newStatus As String, _
                            ByVal newLocation As String, _
                            Optional ByVal setOpened As Boolean = False, _
                            Optional ByVal setDisposed As Boolean = False, _
                            Optional ByVal disposalReason As String = "")
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_CONTAINERS)
    modUtilities.UnprotectSheet ws

    ' BOUNDARY 3: before container mutation (transaction already appended;
    ' caller rollback must restore container + remove the transaction row)
    Call modFaultInjection.FaultAt(modFaultInjection.FAULT_BEFORE_CONTAINER_MUTATION)

    SetCellValue lo, rowNum, COL_STATUS, newStatus
    If Len(newLocation) > 0 Then
        SetCellValue lo, rowNum, COL_STORAGE_LOCATION_ID, newLocation
    ElseIf newStatus = STATUS_IN_USE Or newStatus = STATUS_DISPOSED Or newStatus = STATUS_MISSING Then
        ' Taken/Disposed/Missing leaves storage; clear location.
        SetCellValue lo, rowNum, COL_STORAGE_LOCATION_ID, ""
    End If
    If setOpened Then
        If IsEmpty(CellValue(lo, rowNum, COL_OPENED_DATE)) Then
            SetCellValue lo, rowNum, COL_OPENED_DATE, modUtilities.GetNow()
        End If
    End If
    If setDisposed Then
        SetCellValue lo, rowNum, COL_DISPOSAL_DATE, modUtilities.GetNow()
        SetCellValue lo, rowNum, COL_DISPOSAL_REASON, disposalReason
    End If

    ' BOUNDARY 4: during/after partial container mutation (all fields written;
    ' caller rollback must restore the container to its exact pre-state)
    Call modFaultInjection.FaultAt(modFaultInjection.FAULT_DURING_CONTAINER_MUTATION)

    modUtilities.ProtectSheet ws
End Sub

Public Sub RollbackState(ByRef saved As ContainerState)
    ' Restore all mutated container fields from the captured pre-state.
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_CONTAINERS)
    modUtilities.UnprotectSheet ws

    SetCellValue lo, saved.Row, COL_STATUS, saved.status
    SetCellValue lo, saved.Row, COL_STORAGE_LOCATION_ID, saved.StorageLocationID
    SetCellValue lo, saved.Row, COL_OPENED_DATE, saved.OpenedDate
    SetCellValue lo, saved.Row, COL_DISPOSAL_DATE, saved.DisposalDate
    SetCellValue lo, saved.Row, COL_DISPOSAL_REASON, saved.disposalReason
    SetCellValue lo, saved.Row, COL_NOTES, saved.notes

    modUtilities.ProtectSheet ws
End Sub

' ------------------------------------------------------------------ new container (receive)
Public Function AddContainer(ByVal barcode As String, ByVal productID As String, _
                             ByVal lot As String, ByVal expiryDate As Variant, _
                             ByVal retestDate As Variant, ByVal locationID As String, _
                             ByVal notes As String) As String
    ' Creates a new container row (Status = Available). Returns ContainerID.
    ' Self-cleaning: if a fault fires after the row is allocated, the row
    ' created by THIS uncommitted call is removed before raising, so no
    ' partial container can survive even if the caller has no handle.
    On Error GoTo addFail

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_CONTAINERS)
    modUtilities.UnprotectSheet ws

    Dim cid As String
    cid = NextContainerID()

    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    Dim r As Long
    r = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.count - 1

    ' Barcode / ContainerID / ProductID must be stored as TEXT (frozen
    ' contract: barcode \d{7} text). Without a text number format, Excel
    ' coerces "0000001" to the number 1, which breaks text lookup.
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_CONTAINER_ID)).NumberFormat = "@"
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_BARCODE)).NumberFormat = "@"
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_PRODUCT_ID)).NumberFormat = "@"
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_CONTAINER_ID)).Value2 = cid
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_BARCODE)).Value2 = barcode
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_PRODUCT_ID)).Value2 = productID
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_BATCH_LOT)).Value2 = lot
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_EXPIRY_DATE)).Value2 = IIf(IsBlankOrEmpty(expiryDate), Empty, expiryDate)
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_RETEST_DATE)).Value2 = IIf(IsBlankOrEmpty(retestDate), Empty, retestDate)
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_DATE_RECEIVED)).Value2 = modUtilities.GetNow()
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_STORAGE_LOCATION_ID)).Value2 = locationID
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_STATUS)).Value2 = STATUS_AVAILABLE
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_OPENED_DATE)).Value2 = Empty
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_DISPOSAL_DATE)).Value2 = Empty
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_DISPOSAL_REASON)).Value2 = Empty
    ws.Cells(r, modBarcodeLookup.ColumnIndex(lo, COL_NOTES)).Value2 = notes

    ' BOUNDARY 4 (receive variant): during/after partial container creation.
    ' Self-clean: delete the row we just created, then raise.
    If modFaultInjection.FaultAt(modFaultInjection.FAULT_DURING_CONTAINER_MUTATION) Then
        On Error Resume Next
        lo.ListRows(lo.ListRows.count).Delete
        On Error GoTo addFail
        modUtilities.ProtectSheet ws
        Err.Raise vbObjectError + 2303, "modContainers", _
                  "INJECTED FAULT during container creation (row removed)"
    End If

    modUtilities.ProtectSheet ws
    AddContainer = cid
    Exit Function

addFail:
    On Error Resume Next
    If Not lo Is Nothing Then
        If Not ws Is Nothing Then modUtilities.UnprotectSheet ws
        If Len(cid) > 0 Then
            Dim rowIdx As Long
            rowIdx = modBarcodeLookup.FindContainerRowByID(cid)
            If rowIdx > 0 Then lo.ListRows(rowIdx).Delete
        End If
        If Not ws Is Nothing Then modUtilities.ProtectSheet ws
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 2304, "modContainers", _
              "Container creation failed and was rolled back: " & Err.Description
End Function

' ------------------------------------------------------------------ helpers
Private Function CellValue(ByRef lo As ListObject, ByVal rowNum As Long, ByVal colName As String) As Variant
    Dim c As Long
    c = modBarcodeLookup.ColumnIndex(lo, colName)
    If c <= 0 Then Exit Function
    CellValue = lo.DataBodyRange.Cells(rowNum, c).Value2
End Function

Private Sub SetCellValue(ByRef lo As ListObject, ByVal rowNum As Long, ByVal colName As String, ByVal value As Variant)
    Dim c As Long
    c = modBarcodeLookup.ColumnIndex(lo, colName)
    If c <= 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowNum, c).Value2 = value
End Sub
