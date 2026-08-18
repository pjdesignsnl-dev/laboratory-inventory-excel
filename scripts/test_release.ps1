# Release verification: open LabInventory_v1.0.0.xlsm, run contract check +
# core sweep, capture output.
param([string]$TestName = "modTestHooks.Test_RunCore")
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0.xlsm"
$OutFile = "$Root\evidence\vba\test-hooks-output.txt"
$LogFile = "$Root\evidence\vba\release-test-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $prevEvents = $excel.EnableEvents
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened release workbook (events disabled)")
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run($TestName)
    $sw.Stop()
    $lines.Add("Run " + $TestName + " in " + $sw.ElapsedMilliseconds + " ms")
    Start-Sleep -Seconds 1
    if (Test-Path $OutFile) { foreach ($l in Get-Content $OutFile) { $lines.Add("OUT: " + $l) } }
    else { $lines.Add("out file NOT created") }
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    try { $excel.EnableEvents = $prevEvents } catch {}
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
