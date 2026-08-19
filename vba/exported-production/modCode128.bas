Attribute VB_Name = "modCode128"
Option Explicit

' ============================================================================
' modCode128 - Code 128 label preparation support (Stage 13)
' ============================================================================
' Printer-independent: produces the Code 128 pattern string for a barcode so
' any font/label solution can render it. Code 128B (full ASCII, no check on
' characters) with a computed modulo-103 check digit.
' ============================================================================

Public Function Code128Pattern(ByVal barcode As String) As String
    ' Returns the raw Code 128B encoding as a string of characters
    ' (Start B + data + check + Stop). A TrueType Code 128 font can render it.
    If Len(barcode) = 0 Then Exit Function
    Dim i As Long
    Dim sum As Long
    Dim weight As Long
    sum = 104  ' Code B start value
    weight = 1
    For i = 1 To Len(barcode)
        Dim ch As Long
        ch = AscW(Mid$(barcode, i, 1))
        ' Code B values: ASCII 32..126 -> 0..94
        If ch < 32 Or ch > 126 Then
            Err.Raise vbObjectError + 2701, "modCode128", _
                      "Barcode contains characters outside Code 128B range."
        End If
        Dim val As Long
        val = ch - 32
        sum = sum + (val * weight)
        weight = weight + 1
    Next i
    Dim check As Long
    check = sum Mod 103
    Code128Pattern = ChrW$(204) & barcode & ChrW$(32 + check) & ChrW$(206)
End Function

Public Function BarcodeLabelText(ByVal barcode As String) As String
    ' Human-readable label text: pattern + the barcode digits underneath.
    BarcodeLabelText = Code128Pattern(barcode) & vbLf & barcode
End Function

Public Function BarcodeIsPrintable(ByVal barcode As String) As Boolean
    On Error Resume Next
    Call Code128Pattern(barcode)
    If Err.Number = 0 Then BarcodeIsPrintable = True
    On Error GoTo 0
End Function
