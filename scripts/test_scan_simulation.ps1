# Scanner-simulation (keyboard-wedge) test: type a barcode into Scan!D7 with
# events ENABLED (simulating a scanner Enter-suffix), let the Worksheet_Change
# event dispatch to modScanInterface, then read the staging + status message.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0-candidate.xlsm"
$OutFile = "$Root\evidence\vba\scan-sim-out.txt"
$LogFile = "$Root\evidence\vba\scan-sim-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # events ENABLED (default) so Worksheet_Change fires
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened with events enabled")
    $ws = $wb.Worksheets("Scan")
    # simulate scanner: write barcode + Enter-suffix into D7
    $ws.Unprotect()
    $ws.Range("D7").Value2 = "0000001"
    $ws.Protect()
    $lines.Add("typed 0000001 into D7")
    Start-Sleep -Seconds 3
    # read status message (D9)
    $ws.Unprotect()
    $msg = $ws.Range("D9").Value2
    $lines.Add("status-msg=[" + $msg + "]")
    # read staging barcode + lookup state
    $staging = $ws.ListObjects.Item("tblScanResults")
    $stRows = $staging.DataBodyRange.Rows.Count
    $lines.Add("staging-rows=" + $stRows)
    if ($stRows -gt 0) {
        $hdr = @{}
        for ($c = 1; $c -le $staging.ListColumns.Count; $c++) {
            $hdr[$staging.ListColumns.Item($c).Name] = $c
        }
        $bc = $staging.DataBodyRange.Cells.Item(1, $hdr["Barcode"]).Value2
        $st = $staging.DataBodyRange.Cells.Item(1, $hdr["LookupState"]).Value2
        $lines.Add("staging-barcode=[" + $bc + "] state=[" + $st + "]")
        # check the container was resolved: ContainerID column
        if ($hdr.ContainsKey("ContainerID")) {
            $cid = $staging.DataBodyRange.Cells.Item(1, $hdr["ContainerID"]).Value2
            $lines.Add("staging-containerid=[" + $cid + "]")
        }
    }
    $ws.Protect()
    [System.IO.File]::WriteAllLines($OutFile, $lines)
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
