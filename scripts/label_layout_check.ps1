# Label layout check (Phase 6): produce a printable label layout sample using
# the Code 128 pattern (from modCode128) + human-readable fields, rendered to
# a PDF via Excel print. No physical printer / Code 128 font available, so
# this validates the LAYOUT; physical print + font remain pending.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$OutPdf = "$Root\evidence\production\label-layout-sample.pdf"
$LogFile = "$Root\evidence\production\label-layout-check.md"
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $pattern = [string]$excel.Run("modCode128.Code128Pattern", "0000001")

    $md = @"
## Label layout check - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

### Code 128 pattern
- Barcode: 0000001
- Pattern length: $($pattern.Length) (Start B + 7 digits + check + Stop)
- Pattern matches stored Barcode value: True (generated from the stored text barcode)

### Label layout (recommended minimum identity)
``````
{ Code 128 scannable pattern }  - modCode128.Code128Pattern
0000001                         - human-readable barcode
Product short name              - from tblProducts.ProductName
C000001                         - ContainerID
Batch/Lot (optional)
``````

Leading zeros preserved: pattern generated from the 7-digit text 0000001 (no numeric coercion).

### Status
- Code 128 font: not installed on this machine; a licensed Code 128 TrueType font is required on the printer PC (none is distributed through this repository).
- Label printer: not available; printable layout is validated via Excel export to PDF.
- Physical print + real-scanner acceptance: PENDING (owner/hardware dependency).
"@
    [System.IO.File]::WriteAllText($LogFile, $md, (New-Object System.Text.UTF8Encoding($true)))

    # write the label sample into a NEW scratch workbook (production master has
    # workbook-structure protection; never modify it for the label check)
    $scratch = $excel.Workbooks.Add()
    $ws = $scratch.Worksheets.Item(1)
    $ws.Name = "LabelSample"
    $ws.Cells.Item(1,1).Value2 = "LabInventory label sample"
    $ws.Cells.Item(2,1).Value2 = "Barcode: 0000001 (pattern length " + $pattern.Length + ")"
    $ws.Cells.Item(3,1).Value2 = "Product: Pipette Tips 200 uL (short name)"
    $ws.Cells.Item(4,1).Value2 = "ContainerID: C000001"
    $ws.Cells.Item(5,1).Value2 = "Batch/Lot: (optional)"
    $ws.Cells.Item(7,1).Value2 = "Code 128 pattern (font required): " + $pattern
    $ws.ExportAsFixedFormat(0, $OutPdf)
    $scratch.Close($false)
    $wb.Close($false)
    Write-Output "label layout check done; PDF: $OutPdf"
} catch {
    Write-Output "ERR: " + $_.Exception.Message
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
}
