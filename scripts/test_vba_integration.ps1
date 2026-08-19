# VBA integration test: rebuild candidate, run Test_RunAll sweep, read results.
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Xlsm = Join-Path $Root "workbook\LabInventory_v1.0-candidate.xlsm"
$Log = Join-Path $Root "evidence\vba\stage2-8-test-log.txt"
$OutFile = Join-Path $Root "evidence\vba\test-hooks-output.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue

function Invoke-Com([scriptblock]$block, [string]$label, [int]$attempts = 20) {
    for ($a = 1; $a -le $attempts; $a++) {
        try { return & $block } catch { if ($a -eq $attempts) { throw }; Start-Sleep -Milliseconds 700 }
    }
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = Invoke-Com { $excel.Workbooks.Open($Xlsm, 0, $false) } "Open"
    $lines.Add("opened candidate")

    # single-shot runs (no retry wrapper around Run)
    $excel.Run("modTestHooks.Test_RunAll")
    $lines.Add("Test_RunAll executed")

    Start-Sleep -Seconds 2
    if (Test-Path $OutFile) {
        $lines.Add("--- test-hooks-output.txt ---")
        foreach ($l in Get-Content $OutFile) { $lines.Add($l) }
    } else {
        $lines.Add("test-hooks-output.txt NOT created")
    }
    $wb.Close($false)
} catch {
    $lines.Add("FATAL: " + $_.Exception.Message)
    if ($wb) { try { $wb.Close($false) } catch {} }
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($Log, $lines)
    Write-Output ($lines -join "`n")
}
