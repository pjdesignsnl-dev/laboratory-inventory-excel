VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Phase B principle 2: validate the runtime contract at startup; fail
    ' closed if the workbook structure has drifted.
    On Error GoTo contractFail

    If modWorkbookContract.ContractValidate(True) Then
        Application.StatusBar = "Laboratory Inventory ready."
        modErrorHandling.LogInfo "Workbook_Open", "Contract validation passed."
    Else
        modWorkbookContract.FailClosed modWorkbookContract.ContractDiagnostics()
        Application.StatusBar = "WORKBOOK CONTRACT VIOLATION - operations disabled."
        Exit Sub
    End If
    Exit Sub

contractFail:
    modWorkbookContract.FailClosed "Unable to validate workbook contract at startup: " & Err.Description
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

