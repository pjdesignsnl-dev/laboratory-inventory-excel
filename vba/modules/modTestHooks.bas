Attribute VB_Name = "modTestHooks"
Option Explicit

' ============================================================================
' modTestHooks - Test hooks for VBA-focused tests
' ============================================================================
' Procedures drive internal routines from outside (Application.Run from COM).
' Results are written to a log file (reliable). Every operation is wrapped in
' error suppression so a failure can never block the COM caller with a dialog.
' Not part of the operator workflow.
' ============================================================================

Private Const TEST_LOG As String = "C:\Users\Q\Documents\laboratory-inventory-excel\evidence\vba\test-hooks-output.txt"

Private Sub LogLine(ByVal s As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open TEST_LOG For Append As #ff
    Print #ff, s
    Close #ff
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 1: contract check
Public Sub Test_ContractCheck()
    On Error Resume Next
    Dim ok As Boolean
    ok = modWorkbookContract.ContractValidate(True)
    LogLine "contract:" & IIf(ok, "OK", "FAIL") & " diag=[" & modWorkbookContract.ContractDiagnostics() & "]"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------ Stage 1: contract drift detection
' Drift is simulated by deleting a critical named range, validating (should
' FAIL with a named-range diagnostic), then re-adding it and validating again
' (should OK). This exercises the fail-closed path without worksheet rename
' (which blocks COM re-entrancy).
Public Sub Test_ContractDrift()
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' capture the RefersTo of rngScanInput so we can restore it exactly
    Dim savedRefersTo As String
    savedRefersTo = wb.Names(RNG_SCAN_INPUT).RefersTo
    wb.Names(RNG_SCAN_INPUT).Delete

    Dim okBefore As Boolean
    okBefore = modWorkbookContract.ContractValidate(True)

    wb.Names.Add Name:=RNG_SCAN_INPUT, RefersTo:=savedRefersTo
    Dim okAfter As Boolean
    okAfter = modWorkbookContract.ContractValidate(True)

    LogLine "drift:before=" & IIf(okBefore, "OK", "FAIL") & " after=" & IIf(okAfter, "OK", "FAIL")
    LogLine "drift-diag=[" & modWorkbookContract.ContractDiagnostics() & "]"
    On Error GoTo 0
End Sub
