# Build LabInventory_v1.0-candidate.xlsm from the approved macro-free workbook,
# importing the source-controlled VBA modules (Stage 1: modConstants +
# modWorkbookContract; extended in later stages).
# Requires: Microsoft Excel 16.0+, AccessVBOM=1 (development, reversible).
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SrcXlsx = Join-Path $Root "workbook\LabInventory_v0.1.xlsx"
$DstXlsm = Join-Path $Root "workbook\LabInventory_v1.0-candidate.xlsm"
$ModDir = Join-Path $Root "vba\modules"
$Log = Join-Path $Root "evidence\vba\build-candidate-log.txt"
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
    $wb = Invoke-Com { $excel.Workbooks.Open($SrcXlsx, 0, $false) } "Open"
    $lines.Add("opened " + $SrcXlsx)

    # SaveAs xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    Invoke-Com { $wb.SaveAs($DstXlsm, 52) } "SaveAs"
    $lines.Add("saved " + $DstXlsm)

    # Import modules
    $modFiles = @(
        "modConstants.bas",
        "modWorkbookContract.bas",
        "modTestHooks.bas"
    )
    $vbProject = $wb.VBProject
    foreach ($mf in $modFiles) {
        $p = Join-Path $ModDir $mf
        Invoke-Com { $vbProject.VBComponents.Import($p) } "Import $mf"
        $lines.Add("imported " + $mf)
    }

    # Compile: Excel compiles the VBA project when the workbook is saved.
    Invoke-Com { $wb.Save() } "Save"
    $lines.Add("saved after import (Excel compiles on save)")

    # module inventory
    $compCount = [int](Invoke-Com { $vbProject.VBComponents.Count } "CompCount")
    $lines.Add("VBComponents count: " + $compCount)
    for ($i = 1; $i -le $compCount; $i++) {
        $c = Invoke-Com { $vbProject.VBComponents.Item($i) } "Comp$i"
        $lines.Add("  " + $c.Type + " " + $c.Name)
    }

    $wb.Close($false)
    $lines.Add("closed")
} catch {
    $lines.Add("FATAL: " + $_.Exception.Message)
    if ($wb) { try { $wb.Close($false) } catch {} }
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($Log, $lines)
    Write-Output ($lines -join "`n")
}
