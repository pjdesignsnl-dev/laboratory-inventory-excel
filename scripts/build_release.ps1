# Final release build: produce workbook/LabInventory_v1.0.0.xlsm from the
# hash-free macro-free base + all source-controlled modules + document-module
# code, compile via save, export every component back for verification.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SrcXlsx = Join-Path $Root "workbook\LabInventory_v0.1.xlsx"
$DstXlsm = Join-Path $Root "workbook\LabInventory_v1.0.0.xlsm"
$ModDir = Join-Path $Root "vba\modules"
$DocDir = Join-Path $Root "vba\docmodules"
$ExportDir = Join-Path $Root "vba\exported-release"
$Log = Join-Path $Root "evidence\vba\build-release-log.txt"
$lines = New-Object System.Collections.Generic.List[string]

New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null

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
    $wb = Invoke-Com { $excel.Workbooks.Open($SrcXlsx, 0, $false) } "Open"
    $lines.Add("opened " + $SrcXlsx)
    Invoke-Com { $wb.SaveAs($DstXlsm, 52) } "SaveAs"
    $lines.Add("saved " + $DstXlsm)

    $vbProject = $wb.VBProject
    $modFiles = @(
        "modConstants.bas", "modWorkbookContract.bas", "modUtilities.bas",
        "modBarcodeLookup.bas", "modValidation.bas", "modTransactions.bas",
        "modContainers.bas", "modReceiving.bas", "modScanInterface.bas",
        "modBackup.bas", "modErrorHandling.bas", "modCode128.bas",
        "modFaultInjection.bas", "modTestHooks.bas"
    )
    foreach ($mf in $modFiles) {
        $p = Join-Path $ModDir $mf
        Invoke-Com { $vbProject.VBComponents.Import($p) } "Import $mf"
        $lines.Add("imported " + $mf)
    }

    $twCode = [System.IO.File]::ReadAllText((Join-Path $DocDir "ThisWorkbook.cls"), [System.Text.Encoding]::UTF8)
    $scanCode = [System.IO.File]::ReadAllText((Join-Path $DocDir "Scan.cls"), [System.Text.Encoding]::UTF8)
    $scanCodeName = [string](Invoke-Com { $wb.Worksheets("Scan").CodeName } "ScanCodeName")
    $lines.Add("Scan worksheet CodeName: " + $scanCodeName)

    foreach ($entry in @(@("ThisWorkbook", $twCode), @($scanCodeName, $scanCode))) {
        $targetName = $entry[0]
        $codeText = $entry[1]
        $c = $null
        for ($i = 1; $i -le [int](Invoke-Com { $vbProject.VBComponents.Count } "CC"); $i++) {
            $n = [string](Invoke-Com { $vbProject.VBComponents.Item($i).Name } "N$i")
            if ($n -eq $targetName) { $c = $vbProject.VBComponents.Item($i); break }
        }
        if ($null -ne $c) {
            $cm = $c.CodeModule
            Invoke-Com { $cm.DeleteLines(1, $cm.CountOfLines) } "Clear $targetName"
            $codeLines = $codeText -split "`r?`n"
            $filtered = @($codeLines | Where-Object { $_ -notmatch '^Attribute VB_Name' })
            $body = ($filtered -join "`n")
            Invoke-Com { $cm.AddFromString($body) } "Inject $targetName"
            $lines.Add("injected document module code: " + $targetName)
        } else {
            $lines.Add("WARN: document component " + $targetName + " not found")
        }
    }

    Invoke-Com { $wb.Save() } "Save"
    $lines.Add("saved after import (Excel compiles on save)")

    $cc = [int](Invoke-Com { $vbProject.VBComponents.Count } "CompCount")
    $lines.Add("VBComponents count: " + $cc)
    for ($i = 1; $i -le $cc; $i++) {
        $c = Invoke-Com { $vbProject.VBComponents.Item($i) } "Comp$i"
        $t = [int](Invoke-Com { $c.Type } "CompT$i")
        $n = [string](Invoke-Com { $c.Name } "CompN$i")
        $ext = switch ($t) { 1 { "bas" } 100 { "cls" } 2 { "cls" } 3 { "frm" } default { "bin" } }
        $out = Join-Path $ExportDir ($n + "." + $ext)
        Invoke-Com { $c.Export($out) } "Export $n"
        $lines.Add("exported " + $n + "." + $ext)
    }

    try {
        Invoke-Com { $excel.VBE.MainWindow.Visible = $false } "VbeHide"
        $lines.Add("VBE hidden before final save")
    } catch { $lines.Add("VBE hide failed: " + $_.Exception.Message) }
    Invoke-Com { $wb.Save() } "SaveFinal"
    $lines.Add("final save")

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
