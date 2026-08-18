# Run a single named test hook from the candidate, capture output.
# No .VBE access (that can hang); relies on the build having hidden the VBE.
param([string]$TestName = "modTestHooks.Test_RunCore")
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Xlsm = Join-Path $Root "workbook\LabInventory_v1.0-candidate.xlsm"
$OutFile = Join-Path $Root "evidence\vba\test-hooks-output.txt"
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
    $lines.Add("opened (events disabled)")
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
    [System.IO.File]::WriteAllLines((Join-Path $Root "evidence\vba\run-single-log.txt"), $lines)
    Write-Output ($lines -join "`n")
}
