# Scanner-simulation v2: test multiple event scenarios (unknown, invalid, empty).
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0-candidate.xlsm"
$OutFile = "$Root\evidence\vba\scan-sim2-out.txt"
$LogFile = "$Root\evidence\vba\scan-sim2-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened with events enabled")
    $ws = $wb.Worksheets("Scan")
    function Sim-Scan([string]$bc) {
        $ws.Unprotect()
        $ws.Range("D7").Value2 = $bc
        $ws.Protect()
        Start-Sleep -Seconds 2
        $ws.Unprotect()
        $m = $ws.Range("D9").Value2
        $staging = $ws.ListObjects.Item("tblScanResults")
        $st = $staging.DataBodyRange.Cells.Item(1, 2).Value2  # LookupState col 2
        $ws.Protect()
        return "$m|$st"
    }
    $r1 = Sim-Scan "9999999"
    $lines.Add("unknown-scan=[" + $r1 + "]")
    $r2 = Sim-Scan "abc1234"
    $lines.Add("invalid-scan=[" + $r2 + "]")
    $r3 = Sim-Scan ""
    $lines.Add("empty-scan=[" + $r3 + "]")
    $r4 = Sim-Scan "0000021"
    $lines.Add("expired-avail-scan=[" + $r4 + "]")
    [System.IO.File]::WriteAllLines($OutFile, $lines)
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
