VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Scan worksheet module - DISPATCH ONLY (Phase B principle 1).
' Captures barcode input and dispatches to the service routine in
' modScanInterface. No business rules live here.

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.count > 1 Then Exit Sub
    If Target.Address(False, False) <> "D7" Then Exit Sub
    If modScanInterface.IsBusy Then Exit Sub
    modScanInterface.HandleScannedBarcode Target.Value2
    On Error GoTo 0
End Sub

