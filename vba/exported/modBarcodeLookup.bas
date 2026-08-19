Attribute VB_Name = "modBarcodeLookup"
Option Explicit

' ============================================================================
' modBarcodeLookup - Barcode lookup and duplicate detection (Stage 2)
' ============================================================================
' Resolves a scanned barcode to its Container row. Uses a pure-VBA dictionary
' built from the frozen tblContainers[Barcode] column (no worksheet-function
' calls), which is deterministic, fast for thousands of rows, and immune to
' typecoercion. Duplicate barcodes are detected explicitly.
' ============================================================================

' Lookup results as Long + named constants (module-qualified Enum types in
' Run-invoked procedures can deadlock VBA under COM; constants avoid that).
Public Const LR_FOUND As Long = 0
Public Const LR_UNKNOWN As Long = 1
Public Const LR_DUPLICATE As Long = 2
Public Const LR_EMPTY As Long = 3
Public Const LR_INVALID As Long = 4

' ------------------------------------------------------------------ public API
' NOTE: no ByRef parameters in the public lookup API - a ByRef parameter on a
' module Function invoked from a Run-driven procedure can deadlock VBA/COM in
' some environments. Callers use FindBarcodeRow + BarcodeCount directly.

Public Function LookupState(ByVal barcode As String) As Long
    On Error GoTo handler

    barcode = modUtilities.NormalizeBarcode(barcode)
    If Len(barcode) = 0 Then
        LookupState = LR_EMPTY
        Exit Function
    End If
    If Not modUtilities.IsText7DigitBarcode(barcode) Then
        LookupState = LR_INVALID
        Exit Function
    End If

    Dim idx As Long
    idx = BarcodeRowIndex(barcode)
    If idx = 0 Then
        LookupState = LR_UNKNOWN
        Exit Function
    End If
    If BarcodeCount(barcode) > 1 Then
        LookupState = LR_DUPLICATE
        Exit Function
    End If
    LookupState = LR_FOUND
    Exit Function

handler:
    LookupState = LR_INVALID
End Function

Public Function FindBarcodeRow(ByVal barcode As String) As Long
    ' Returns the data-body row (1-based) of the first match, or 0.
    On Error Resume Next
    FindBarcodeRow = BarcodeRowIndex(modUtilities.NormalizeBarcode(barcode))
    On Error GoTo 0
End Function

Public Function BarcodeCount(ByVal barcode As String) As Long
    On Error Resume Next
    BarcodeCount = BarcodeRowIndexCount(barcode)
    On Error GoTo 0
End Function

Public Function FindContainerRowByID(ByVal containerID As String) As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = ColumnIndex(lo, COL_CONTAINER_ID)
    If colIdx <= 0 Then Exit Function
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        If CStr(lo.DataBodyRange.Cells(r, colIdx).Value2) = containerID Then
            FindContainerRowByID = r
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ stateless linear scan
' No module-level state (a module-level dynamic array can deadlock VBA/COM in
' some Run-invoked contexts). Each call scans the frozen tblContainers[Barcode]
' column directly: instant for the fixture scale and fine for thousands of rows.
Private Function BarcodeRowIndex(ByVal barcode As String) As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = ColumnIndex(lo, COL_BARCODE)
    If colIdx <= 0 Then Exit Function
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, colIdx).Value2), barcode, vbBinaryCompare) = 0 Then
            BarcodeRowIndex = r
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function

Private Function BarcodeRowIndexCount(ByVal barcode As String) As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = ColumnIndex(lo, COL_BARCODE)
    If colIdx <= 0 Then Exit Function
    Dim r As Long
    Dim c As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, colIdx).Value2), barcode, vbBinaryCompare) = 0 Then c = c + 1
    Next r
    BarcodeRowIndexCount = c
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ staging population
Public Sub PopulateScanStaging(ByVal barcode As String, ByVal rowNum As Long, ByVal result As Long)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_SCAN).ListObjects(TBL_SCAN_RESULTS)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_SCAN)
    modUtilities.UnprotectSheet ws

    Dim r As Long
    r = lo.DataBodyRange.row
    If lo.DataBodyRange.Rows.count = 0 Then
        lo.ListRows.Add
        r = lo.DataBodyRange.row
    End If

    SetCellValue ws, r, "Barcode", barcode
    SetCellValue ws, r, "LookupState", LookupStateText(result)

    If result = LR_FOUND Or result = LR_DUPLICATE Then
        Dim contLo As ListObject
        Set contLo = ThisWorkbook.Worksheets(WS_CONTAINERS).ListObjects(TBL_CONTAINERS)
        SetCellValue ws, r, "ContainerID", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_CONTAINER_ID)).Value2
        SetCellValue ws, r, "ProductID", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_PRODUCT_ID)).Value2
        SetCellValue ws, r, "BatchLotNumber", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_BATCH_LOT)).Value2
        SetCellValue ws, r, "ExpiryDate", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_EXPIRY_DATE)).Value2
        SetCellValue ws, r, "StorageLocationID", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_STORAGE_LOCATION_ID)).Value2
        SetCellValue ws, r, "Status", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_STATUS)).Value2
        SetCellValue ws, r, "OpenedDate", contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_OPENED_DATE)).Value2
        Dim pid As String
        pid = CStr(contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_PRODUCT_ID)).Value2)
        SetCellValue ws, r, "ProductName", productName(pid)
        Dim loc As String
        loc = CStr(contLo.DataBodyRange.Cells(rowNum, ColumnIndex(contLo, COL_STORAGE_LOCATION_ID)).Value2)
        SetCellValue ws, r, "LocationName", LocationName(loc)
    Else
        SetCellValue ws, r, "ContainerID", ""
        SetCellValue ws, r, "ProductID", ""
        SetCellValue ws, r, "ProductName", ""
        SetCellValue ws, r, "BatchLotNumber", ""
        SetCellValue ws, r, "ExpiryDate", ""
        SetCellValue ws, r, "StorageLocationID", ""
        SetCellValue ws, r, "LocationName", ""
        SetCellValue ws, r, "Status", ""
        SetCellValue ws, r, "OpenedDate", ""
    End If

    modUtilities.ProtectSheet ws
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ helpers
Public Function ColumnIndex(ByRef lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            ColumnIndex = i
            Exit Function
        End If
    Next i
    ColumnIndex = 0
End Function

Public Function productName(ByVal productID As String) As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_PRODUCTS).ListObjects(TBL_PRODUCTS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long, nmCol As Long
    idCol = ColumnIndex(lo, COL_PRODUCT_ID)
    nmCol = ColumnIndex(lo, COL_PRODUCT_NAME)
    If idCol <= 0 Or nmCol <= 0 Then Exit Function
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        If CStr(lo.DataBodyRange.Cells(r, idCol).Value2) = productID Then
            productName = CStr(lo.DataBodyRange.Cells(r, nmCol).Value2)
            Exit Function
        End If
    Next r
End Function

Public Function LocationName(ByVal locationID As String) As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(WS_LOCATIONS).ListObjects(TBL_LOCATIONS)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idCol As Long, nmCol As Long
    idCol = ColumnIndex(lo, COL_STORAGE_LOCATION_ID)
    nmCol = ColumnIndex(lo, "LocationName")
    If idCol <= 0 Or nmCol <= 0 Then Exit Function
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        If CStr(lo.DataBodyRange.Cells(r, idCol).Value2) = locationID Then
            LocationName = CStr(lo.DataBodyRange.Cells(r, nmCol).Value2)
            Exit Function
        End If
    Next r
End Function

Private Sub SetCellValue(ByRef ws As Worksheet, ByVal row As Long, ByVal colName As String, ByVal value As Variant)
    Dim hdr As Range
    Set hdr = ws.Rows(loHeaderRow(ws)).Find(What:=colName, LookAt:=xlWhole)
    If hdr Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = ws.Cells(row, hdr.Column)
    ' Barcode and ID cells must be TEXT (leading zeros preserved, no coercion)
    If colName = "Barcode" Or colName = "ContainerID" Or colName = "ProductID" Then
        cell.NumberFormat = "@"
    End If
    cell.Value2 = value
End Sub

Private Function loHeaderRow(ByRef ws As Worksheet) As Long
    loHeaderRow = ws.ListObjects(TBL_SCAN_RESULTS).HeaderRowRange.row
End Function

Private Function LookupStateText(ByVal result As Long) As String
    Select Case result
        Case LR_FOUND: LookupStateText = "FOUND"
        Case LR_UNKNOWN: LookupStateText = "UNKNOWN"
        Case LR_DUPLICATE: LookupStateText = "DUPLICATE"
        Case LR_EMPTY: LookupStateText = "EMPTY"
        Case Else: LookupStateText = "INVALID"
    End Select
End Function
