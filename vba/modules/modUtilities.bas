Attribute VB_Name = "modUtilities"
Option Explicit

' ============================================================================
' modUtilities - Time/user abstraction, application-state safety, helpers
' ============================================================================
' Centralized clock and operator capture (Phase B principles 10 and 5), plus
' small string/date helpers used across modules. No scattered Now/Date.
'
' NOTE: no Public Type / UDT declarations in this module - a Public Type in a
' standard module referenced from a Run-invoked procedure deadlocks VBA/COM in
' some environments. State save/restore uses individual module-level variables.
' ============================================================================

Private m_fixedNow As Date

' ------------------------------------------------------------------ time/user
Public Function GetNow() As Date
    ' Central time source (improves testability; tests can override via
    ' SetNowForTest when a stable clock is required).
    If m_fixedNow > 0 Then
        GetNow = m_fixedNow
    Else
        GetNow = Now
    End If
End Function

Public Sub SetNowForTest(ByVal dt As Date)
    m_fixedNow = dt
End Sub

Public Sub ResetNowForTest()
    m_fixedNow = 0
End Sub

Public Function GetOperator() As String
    ' Windows username capture (reserved Operator column; no login system).
    Dim u As String
    On Error Resume Next
    u = Environ$("USERNAME")
    On Error GoTo 0
    If Len(Trim$(u)) = 0 Then u = OPERATOR_UNKNOWN
    GetOperator = u
End Function

' ------------------------------------------------------------------ barcode/ID helpers
Public Function IsText7DigitBarcode(ByVal v As Variant) As Boolean
    On Error Resume Next
    If VarType(v) <> vbString Then Exit Function
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) <> 7 Then Exit Function
    Dim i As Long
    For i = 1 To 7
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsText7DigitBarcode = True
    On Error GoTo 0
End Function

Public Function NormalizeBarcode(ByVal v As Variant) As String
    ' Barcodes are stored as 7-digit text. If a numeric value was entered
    ' (e.g., scanner without leading zero), re-pad to 7 digits.
    On Error Resume Next
    If VarType(v) = vbString Then
        NormalizeBarcode = Trim$(CStr(v))
    ElseIf IsNumeric(v) Then
        NormalizeBarcode = Format$(CLng(v), FMT_BARCODE)
    Else
        NormalizeBarcode = CStr(v)
    End If
    On Error GoTo 0
End Function

Public Function ParseContainerNum(ByVal containerID As String) As Long
    ' "C000123" -> 123 (mirrors the frozen HelperContainerNum formula).
    On Error Resume Next
    ParseContainerNum = CLng(Mid$(containerID, 2))
    On Error GoTo 0
End Function

' ------------------------------------------------------------------ application state safety
' Individual save/restore (no UDT) so no exotic type reference can deadlock.
Public Sub SaveAppState(ByRef ev As Boolean, ByRef sc As Boolean, ByRef da As Boolean)
    On Error Resume Next
    ev = Application.EnableEvents
    sc = Application.ScreenUpdating
    da = Application.DisplayAlerts
    On Error GoTo 0
End Sub

Public Sub RestoreAppState(ByVal ev As Boolean, ByVal sc As Boolean, ByVal da As Boolean)
    On Error Resume Next
    Application.EnableEvents = ev
    Application.ScreenUpdating = sc
    Application.DisplayAlerts = da
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ sheet protection helpers
Public Sub UnprotectSheet(ByRef ws As Worksheet)
    On Error Resume Next
    If ws.ProtectContents Then ws.Unprotect
    On Error GoTo 0
End Sub

Public Sub ProtectSheet(ByRef ws As Worksheet)
    ' Restore protection with the empty-password deterrent (D-015).
    On Error Resume Next
    If Not ws.ProtectContents Then ws.Protect
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ table helpers
Public Function TableDataRange(ByVal tableName As String) As Range
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(tableSheet(tableName)).ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then Set TableDataRange = lo.DataBodyRange
End Function

Private Function tableSheet(ByVal tableName As String) As String
    ' Map table name -> worksheet (frozen contract).
    Select Case tableName
        Case TBL_PRODUCTS: tableSheet = WS_PRODUCTS
        Case TBL_CONTAINERS: tableSheet = WS_CONTAINERS
        Case TBL_TRANSACTIONS: tableSheet = WS_TRANSACTIONS
        Case TBL_SUPPLIERS: tableSheet = WS_SUPPLIERS
        Case TBL_LOCATIONS: tableSheet = WS_LOCATIONS
        Case TBL_SETTINGS, TBL_STATUS_LIST, TBL_TXN_TYPE_LIST, TBL_EXPIRY_CLASS_LIST: tableSheet = WS_SETTINGS
        Case TBL_SCAN_RESULTS: tableSheet = WS_SCAN
        Case TBL_RECEIVE_STAGING: tableSheet = WS_RECEIVING
        Case Else: tableSheet = ""
    End Select
End Function

' ------------------------------------------------------------------ misc
Public Function IsBlankOrEmpty(ByVal v As Variant) As Boolean
    If IsNull(v) Then IsBlankOrEmpty = True: Exit Function
    If IsEmpty(v) Then IsBlankOrEmpty = True: Exit Function
    If VarType(v) = vbString And Len(Trim$(CStr(v))) = 0 Then IsBlankOrEmpty = True
End Function

Public Function Coalesce(ByVal v As Variant, ByVal fallback As Variant) As Variant
    If IsBlankOrEmpty(v) Then Coalesce = fallback Else Coalesce = v
End Function
