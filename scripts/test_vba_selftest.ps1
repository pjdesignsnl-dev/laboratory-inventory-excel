# Self-test driver v2: open the candidate with LABINV_SELFTEST=1, wait for the
# self-test (Workbook_Open) to finish and write its log, then read it.
# No .VBE access, no Application.Run - avoids the COM/VBE deadlock entirely.
$ErrorActionPreference = "Continue"
$Xlsm = "C:\Users\Q\Documents\laboratory-inventory-excel\workbook\LabInventory_v1.0-candidate.xlsm"
$OutFile = "C:\Users\Q\Documents\laboratory-inventory-excel\evidence\vba\test-hooks-output.txt"
$LogFile = "C:\Users\Q\Documents\laboratory-inventory-excel\evidence\vba\selftest-run-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue
$excel = $null
try {
    $env:LABINV_SELFTEST = "1"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened; waiting for self-test to complete")
    $deadline = (Get-Date).AddSeconds(90)
    while (-not (Test-Path $OutFile) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 1000
    }
    if (Test-Path $OutFile) {
        $lines.Add("--- test-hooks-output.txt ---")
        foreach ($l in Get-Content $OutFile) { $lines.Add($l) }
    } else {
        $lines.Add("test log NOT created within 90s")
    }
    Start-Sleep -Seconds 3
    try { $wb.Close($false) } catch {}
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    Remove-Item Env:LABINV_SELFTEST -ErrorAction SilentlyContinue
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
