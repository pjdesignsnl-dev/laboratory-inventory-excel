# Verify the candidate xlsm: reopen, compile-check via module code inspection,
# and attempt the runtime contract validator through Application.Run.
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Xlsm = Join-Path $Root "workbook\LabInventory_v1.0-candidate.xlsm"
$Log = Join-Path $Root "evidence\vba\verify-candidate-log.txt"
$lines = New-Object System.Collections.Generic.List[string]

function Invoke-Com([scriptblock]$block, [string]$label, [int]$attempts = 15) {
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
    $lines.Add("reopened " + $Xlsm)

    # inspect module code lengths (sanity: code present)
    $vp = $wb.VBProject
    $cc = [int](Invoke-Com { $vp.VBComponents.Count } "CC")
    $lines.Add("components: " + $cc)
    for ($i = 1; $i -le $cc; $i++) {
        $c = Invoke-Com { $vp.VBComponents.Item($i) } "C$i"
        $n = Invoke-Com { $c.Name } "N$i"
        $t = Invoke-Com { $c.Type } "T$i"
        if ($t -eq 1) {
            $cm = Invoke-Com { $c.CodeModule } "CM$i"
            $cnt = [int](Invoke-Com { $cm.CountOfLines } "CL$i")
            $lines.Add("  module " + $n + " lines=" + $cnt)
        }
    }

    # Try to run the contract validator via Application.Run
    $runResult = $null
    try {
        $runResult = Invoke-Com { $excel.Application.Run("modWorkbookContract.ContractValidate", $true) } "Run"
        $lines.Add("ContractValidate result: " + $runResult)
    } catch {
        $lines.Add("ContractValidate run error: " + $_.Exception.Message)
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
