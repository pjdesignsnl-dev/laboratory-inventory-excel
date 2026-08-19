Attribute VB_Name = "modErrorHandling"
Option Explicit

' ============================================================================
' modErrorHandling - Error classification, operator messaging, diagnostics
' ============================================================================
' Distinguishes blocking validation error / confirmation warning /
' informational message / unexpected runtime error. Operator-facing messages
' are clear; raw internals go to the diagnostics log only.
' ============================================================================

Private Const DIAG_LOG As String = "LabInventory_diagnostics.log"

Public Enum ErrClass
    ecBlocking = 1
    ecConfirm = 2
    ecInfo = 3
    ecRuntime = 4
End Enum

Public Sub HandleError(ByVal errNumber As Long, ByVal errSource As String, _
                       ByVal errDescription As String, ByVal context As String, _
                       Optional ByVal errClass As ErrClass = ecRuntime)
    LogDiagnostics errNumber, errSource, errDescription, context
    Dim msg As String
    Select Case errClass
        Case ecBlocking: msg = "Blocked: "
        Case ecConfirm: msg = "Please confirm: "
        Case ecInfo: msg = ""
        Case Else: msg = "An unexpected error occurred: "
    End Select
    msg = msg & OperatorMessage(errNumber, errDescription)
    MsgBox msg, IIf(errClass = ecRuntime, vbExclamation, vbInformation), "Laboratory Inventory"
End Sub

Public Function OperatorMessage(ByVal errNumber As Long, ByVal errDescription As String) As String
    ' Map known error codes to operator-friendly text; fall back to the
    ' description (which we keep operator-safe by construction).
    Select Case errNumber
        Case vbObjectError + 2001: OperatorMessage = "Workbook structure does not match the approved contract."
        Case vbObjectError + 2101: OperatorMessage = "Internal error: barcode column missing."
        Case vbObjectError + 2201: OperatorMessage = "Internal error: transaction ID column missing."
        Case vbObjectError + 2202: OperatorMessage = "Transaction could not be recorded correctly."
        Case vbObjectError + 2301: OperatorMessage = "Internal error: container ID column missing."
        Case vbObjectError + 2302: OperatorMessage = "Internal error: barcode number column missing."
        Case vbObjectError + 2401: OperatorMessage = errDescription
        Case vbObjectError + 2402, vbObjectError + 2403: OperatorMessage = "Receive could not be completed correctly."
        Case vbObjectError + 2404: OperatorMessage = "Receive failed and was rolled back. No partial data was saved."
        Case vbObjectError + 2405: OperatorMessage = errDescription
        Case vbObjectError + 2406: OperatorMessage = "Batch receive failed and was rolled back. No partial data was saved."
        Case vbObjectError + 2501, vbObjectError + 2502: OperatorMessage = "The action could not be completed correctly."
        Case vbObjectError + 2601, vbObjectError + 2602: OperatorMessage = "Backup failed. The operation was not continued."
        Case Else: OperatorMessage = errDescription
    End Select
End Function

Private Sub LogDiagnostics(ByVal errNumber As Long, ByVal errSource As String, _
                           ByVal errDescription As String, ByVal context As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open ThisWorkbook.Path & Application.PathSeparator & DIAG_LOG For Append As #ff
    Print #ff, Format$(modUtilities.GetNow(), "yyyy-mm-dd hh:nn:ss") & _
                " [num=" & errNumber & " src=" & errSource & " ctx=" & context & "] " & errDescription
    Close #ff
    On Error GoTo 0
End Sub

Public Sub LogInfo(ByVal context As String, ByVal text As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open ThisWorkbook.Path & Application.PathSeparator & DIAG_LOG For Append As #ff
    Print #ff, Format$(modUtilities.GetNow(), "yyyy-mm-dd hh:nn:ss") & " [info ctx=" & context & "] " & text
    Close #ff
    On Error GoTo 0
End Sub
