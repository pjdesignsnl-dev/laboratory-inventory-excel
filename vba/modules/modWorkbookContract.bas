Attribute VB_Name = "modWorkbookContract"
Option Explicit

' ============================================================================
' modWorkbookContract - Runtime contract validator (Phase B principle 2)
' ============================================================================
' Before allowing any inventory mutation, validate that the workbook structure
' exactly matches the frozen contract. If the structure has drifted, fail
' closed and present a clear diagnostic. No partial operation is permitted.
' ============================================================================

Private m_lastDiagnostics As String
Private m_validated As Boolean
Private m_valid As Boolean
Private m_lastStep As String

' ------------------------------------------------------------------ public API
Public Function ContractValidate(Optional ByVal force As Boolean = False) As Boolean
    If m_validated And Not force Then
        ContractValidate = m_valid
        Exit Function
    End If

    On Error GoTo validateError

    Dim diag As String
    Dim issues As New Collection
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' 1. Worksheets (exact names, exact count)
    m_lastStep = "worksheets"
    CheckWorksheet wb, WS_DASHBOARD, issues
    CheckWorksheet wb, WS_SCAN, issues
    CheckWorksheet wb, WS_RECEIVING, issues
    CheckWorksheet wb, WS_PRODUCTS, issues
    CheckWorksheet wb, WS_CONTAINERS, issues
    CheckWorksheet wb, WS_TRANSACTIONS, issues
    CheckWorksheet wb, WS_SUPPLIERS, issues
    CheckWorksheet wb, WS_LOCATIONS, issues
    CheckWorksheet wb, WS_SETTINGS, issues
    If wb.Worksheets.Count <> 9 Then
        issues.Add "worksheet count: expected 9, found " & wb.Worksheets.Count
    End If

    ' 2. ListObjects on the correct worksheets
    m_lastStep = "tables"
    CheckTable wb, WS_SCAN, TBL_SCAN_RESULTS, issues
    CheckTable wb, WS_RECEIVING, TBL_RECEIVE_STAGING, issues
    CheckTable wb, WS_PRODUCTS, TBL_PRODUCTS, issues
    CheckTable wb, WS_CONTAINERS, TBL_CONTAINERS, issues
    CheckTable wb, WS_TRANSACTIONS, TBL_TRANSACTIONS, issues
    CheckTable wb, WS_SUPPLIERS, TBL_SUPPLIERS, issues
    CheckTable wb, WS_LOCATIONS, TBL_LOCATIONS, issues
    CheckTable wb, WS_SETTINGS, TBL_SETTINGS, issues
    CheckTable wb, WS_SETTINGS, TBL_STATUS_LIST, issues
    CheckTable wb, WS_SETTINGS, TBL_TXN_TYPE_LIST, issues
    CheckTable wb, WS_SETTINGS, TBL_EXPIRY_CLASS_LIST, issues

    ' 3. Critical columns present on the main tables
    m_lastStep = "columns"
    CheckColumns wb, WS_CONTAINERS, TBL_CONTAINERS, _
        Array(COL_CONTAINER_ID, COL_BARCODE, COL_PRODUCT_ID, COL_BATCH_LOT, _
              COL_EXPIRY_DATE, COL_RETEST_DATE, COL_DATE_RECEIVED, _
              COL_STORAGE_LOCATION_ID, COL_STATUS, COL_OPENED_DATE, _
              COL_DISPOSAL_DATE, COL_DISPOSAL_REASON, COL_NOTES, _
              COL_HELPER_CONTAINER_NUM, COL_HELPER_BARCODE_NUM), issues
    CheckColumns wb, WS_TRANSACTIONS, TBL_TRANSACTIONS, _
        Array(COL_TRANSACTION_ID, COL_TIMESTAMP, COL_OPERATOR, COL_BARCODE, _
              COL_CONTAINER_ID, COL_PRODUCT_ID, COL_PRODUCT_NAME, COL_TXN_TYPE, _
              COL_PREVIOUS_STATUS, COL_NEW_STATUS, COL_PREVIOUS_LOCATION, _
              COL_NEW_LOCATION, COL_BATCH_LOT, COL_REASON, COL_REFERENCE, COL_NOTES), issues
    CheckColumns wb, WS_PRODUCTS, TBL_PRODUCTS, _
        Array(COL_PRODUCT_ID, COL_PRODUCT_NAME, COL_SUPPLIER_ID, COL_MIN_STOCK, _
              COL_TARGET_STOCK, COL_REORDER_QTY, COL_ACTIVE), issues

    ' 4. Critical named ranges resolve
    m_lastStep = "names"
    CheckName wb, RNG_SCAN_INPUT, issues
    CheckName wb, RNG_RECV_PRODUCT_ID, issues
    CheckName wb, RNG_RECV_NEXT_CONTAINER_ID, issues
    CheckName wb, RNG_RECV_NEXT_BARCODE, issues
    CheckName wb, NAME_DEFAULT_LOCATION, issues
    CheckName wb, NAME_DEFAULT_NEW_STATUS, issues
    CheckName wb, NAME_EXPIRY_30, issues
    CheckName wb, NAME_EXPIRY_60, issues
    CheckName wb, NAME_EXPIRY_90, issues

    ' 5. Controlled statuses exactly the frozen 6-value set
    m_lastStep = "statuslist"
    CheckStatusList wb, issues
    ' 6. Controlled transaction types exactly the frozen 9-value set
    m_lastStep = "txntypelist"
    CheckTransactionTypeList wb, issues

    m_lastDiagnostics = JoinCollection(issues)
    m_valid = (issues.Count = 0)
    m_validated = True
    ContractValidate = m_valid
    On Error GoTo 0
    Exit Function

validateError:
    ' Record the unexpected error so the caller sees a diagnostic instead of
    ' a silent FAIL with empty diagnostics.
    m_lastDiagnostics = "validator internal error " & Err.Number & " at step '" & m_lastStep & "': " & Err.Description
    m_valid = False
    m_validated = True
    ContractValidate = False
    On Error GoTo 0
End Function

Public Function ContractDiagnostics() As String
    ContractDiagnostics = m_lastDiagnostics
End Function

Public Function ContractIsValid() As Boolean
    If Not m_validated Then
        Call ContractValidate
    End If
    ContractIsValid = m_valid
End Function

Public Sub FailClosed(ByVal message As String)
    Dim msg As String
    msg = "LABORATORY INVENTORY - WORKBOOK CONTRACT VIOLATION" & vbCrLf & vbCrLf
    msg = msg & "The workbook structure does not match the frozen contract. " _
             & "Inventory operations are DISABLED to protect data integrity." _
             & vbCrLf & vbCrLf
    msg = msg & message & vbCrLf & vbCrLf
    msg = msg & "Do not continue. Report this diagnostic to the administrator."
    MsgBox msg, vbCritical, "Contract Violation - Operations Disabled"
End Sub

Public Sub EnforceContract()
    ' Fail-closed entry point used before any mutation.
    If Not ContractIsValid() Then
        FailClosed ContractDiagnostics()
        Err.Raise vbObjectError + 2001, "modWorkbookContract", _
                  "Contract validation failed; operations disabled."
    End If
End Sub

' ------------------------------------------------------------------ internals
Private Sub CheckWorksheet(ByVal wb As Workbook, ByVal name As String, ByRef issues As Collection)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        issues.Add "worksheet '" & name & "' missing"
    End If
End Sub

Private Sub CheckTable(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String, ByRef issues As Collection)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        issues.Add "table '" & tableName & "' - worksheet '" & sheetName & "' missing"
        Exit Sub
    End If
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        issues.Add "ListObject '" & tableName & "' missing on '" & sheetName & "'"
    End If
End Sub

Private Sub CheckColumns(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String, ByVal colNames As Variant, ByRef issues As Collection)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    ' Build a set of the actual column names (header text must equal the name)
    ' using a native Collection (no external COM objects).
    Dim actual As Collection
    Set actual = New Collection
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        Dim cn As String
        cn = lo.ListColumns(i).Name
        On Error Resume Next
        actual.Add cn
        On Error GoTo 0
    Next i

    Dim j As Long
    For j = LBound(colNames) To UBound(colNames)
        Dim want As String
        want = CStr(colNames(j))
        If Not CollectionHas(actual, want) Then
            issues.Add "table '" & tableName & "' missing column '" & want & "'"
        End If
    Next j
End Sub

Private Sub CheckName(ByVal wb As Workbook, ByVal name As String, ByRef issues As Collection)
    ' Check the name exists without materializing its range (whole-column
    ' list names like Settings!$A:$A must never be resolved to a Range).
    Dim nm As Name
    On Error Resume Next
    Set nm = wb.Names(name)
    On Error GoTo 0
    If nm Is Nothing Then
        issues.Add "named range '" & name & "' missing"
    End If
End Sub

Private Sub CheckStatusList(ByVal wb As Workbook, ByRef issues As Collection)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(WS_SETTINGS)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(TBL_STATUS_LIST)
    On Error GoTo 0
    If lo Is Nothing Then
        issues.Add "tblStatusList missing; cannot validate statuses"
        Exit Sub
    End If
    Dim values As Collection
    Set values = New Collection
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        On Error Resume Next
        values.Add CStr(lo.ListRows(i).Range.Cells(1, 1).Value2)
        On Error GoTo 0
    Next i
    Dim expected As Variant
    expected = Array(STATUS_AVAILABLE, STATUS_IN_USE, STATUS_EXPIRED, _
                     STATUS_DAMAGED, STATUS_DISPOSED, STATUS_MISSING)
    Dim j As Long
    For j = LBound(expected) To UBound(expected)
        If Not CollectionHas(values, CStr(expected(j))) Then
            issues.Add "tblStatusList missing status '" & CStr(expected(j)) & "'"
        End If
    Next j
    If values.Count <> 6 Then
        issues.Add "tblStatusList has " & values.Count & " statuses; frozen contract requires 6"
    End If
End Sub

Private Sub CheckTransactionTypeList(ByVal wb As Workbook, ByRef issues As Collection)
    On Error GoTo dbgErr
    Dim ws As Worksheet
    Set ws = wb.Worksheets(WS_SETTINGS)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(TBL_TXN_TYPE_LIST)
    On Error GoTo 0
    If lo Is Nothing Then
        issues.Add "tblTransactionTypeList missing; cannot validate transaction types"
        Exit Sub
    End If
    Dim values As Collection
    Set values = New Collection
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        On Error Resume Next
        values.Add CStr(lo.ListRows(i).Range.Cells(1, 1).Value2)
        On Error GoTo 0
    Next i
    Dim expected As Variant
    expected = Array(TXN_RECEIVE, TXN_TAKE_OPEN, TXN_RETURN, TXN_TRANSFER, _
                     TXN_DISPOSE, TXN_MARK_EXPIRED, TXN_MARK_DAMAGED, _
                     TXN_MARK_MISSING, TXN_ADJUSTMENT)
    Dim j As Long
    For j = LBound(expected) To UBound(expected)
        If Not CollectionHas(values, CStr(expected(j))) Then
            issues.Add "tblTransactionTypeList missing type '" & CStr(expected(j)) & "'"
        End If
    Next j
    If values.Count <> 9 Then
        issues.Add "tblTransactionTypeList has " & values.Count & " types; frozen contract requires 9"
    End If
    Exit Sub
dbgErr:
    issues.Add "txntypelist-internal-error " & Err.Number & ": " & Err.Description
End Sub

Private Function CollectionHas(ByRef c As Collection, ByVal key As String) As Boolean
    ' Iterate items and compare to the key value. No key-access error possible.
    Dim v As Variant
    For Each v In c
        If CStr(v) = key Then
            CollectionHas = True
            Exit Function
        End If
    Next v
    CollectionHas = False
End Function

Private Function JoinCollection(ByRef c As Collection) As String
    If c.Count = 0 Then
        JoinCollection = ""
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(1 To c.Count)
    Dim i As Long
    For i = 1 To c.Count
        parts(i) = CStr(c(i))
    Next i
    JoinCollection = Join(parts, "; ")
End Function
