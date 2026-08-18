Attribute VB_Name = "Scan"
Option Explicit

' Scan worksheet module - DISPATCH ONLY (Phase B principle 1).
' Captures barcode input and dispatches to the service routine in
' modScanInterface. No business rules live here.

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> "D7" Then Exit Sub
    If modScanInterface.IsBusy Then Exit Sub
    modScanInterface.HandleScannedBarcode Target.Value2
    On Error GoTo 0
End Sub
