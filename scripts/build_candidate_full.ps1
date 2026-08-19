# Full candidate build: imports all source-controlled modules + document-module
# code into the candidate .xlsm, compiles (Excel compiles on save), and exports
# every component back to vba/ for source-control verification.
# Requires: Microsoft Excel 16.0+, AccessVBOM=1 (development, reversible).
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SrcXlsx = Join-Path $Root "workbook\LabInventory_v0.1.xlsx"
$DstXlsm = Join-Path $Root "workbook\LabInventory_v1.0-candidate.xlsm"
$ModDir = Join-Path $Root "vba\modules"
$DocDir = Join-Path $Root "vba\docmodules"
$ExportDir = Join-Path $Root "vba\exported"
$Log = Join-Path $Root "evidence\vba\build-candidate-log.txt"
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

    # SaveAs xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    Invoke-Com { $wb.SaveAs($DstXlsm, 52) } "SaveAs"
    $lines.Add("saved " + $DstXlsm)

    $vbProject = $wb.VBProject

    # Import standard modules
    $modFiles = @(
        "modConstants.bas",
        "modWorkbookContract.bas",
        "modUtilities.bas",
        "modBarcodeLookup.bas",
        "modValidation.bas",
        "modTransactions.bas",
        "modContainers.bas",
        "modReceiving.bas",
        "modScanInterface.bas",
        "modBackup.bas",
        "modErrorHandling.bas",
        "modCode128.bas",
        "modFaultInjection.bas",
        "modTestHooks.bas"
    )
    foreach ($mf in $modFiles) {
        $p = Join-Path $ModDir $mf
        Invoke-Com { $vbProject.VBComponents.Import($p) } "Import $mf"
        $lines.Add("imported " + $mf)
    }

    # Inject document-module code into ThisWorkbook and the Scan worksheet
    # module. The Scan worksheet's CodeName is SheetN; resolve via
    # Worksheets("Scan").CodeName.
    $twCode = [System.IO.File]::ReadAllText((Join-Path $DocDir "ThisWorkbook.cls"), [System.Text.Encoding]::UTF8)
    $scanCode = [System.IO.File]::ReadAllText((Join-Path $DocDir "Scan.cls"), [System.Text.Encoding]::UTF8)
    $scanCodeName = [string](Invoke-Com { $wb.Worksheets("Scan").CodeName } "ScanCodeName")
    $lines.Add("Scan worksheet CodeName: " + $scanCodeName)

    foreach ($entry in @(@("ThisWorkbook", $twCode), @($scanCodeName, $scanCode))) {
        $targetName = $entry[0]
        $codeText = $entry[1]
        $c = $null
        for ($i = 1; $i -le [int](Invoke-Com { $vbProject.VBComponents.Count } "CC"); $i++) {
            $t = [int](Invoke-Com { $vbProject.VBComponents.Item($i).Type } "T$i")
            $n = [string](Invoke-Com { $vbProject.VBComponents.Item($i).Name } "N$i")
            if ($n -eq $targetName) { $c = $vbProject.VBComponents.Item($i); break }
        }
        if ($null -ne $c) {
            $cm = $c.CodeModule
            Invoke-Com { $cm.DeleteLines(1, $cm.CountOfLines) } "Clear $targetName"
            # strip the Attribute VB_Name line (not allowed in AddFromString for doc modules)
            $codeLines = $codeText -split "`r?`n"
            $filtered = @($codeLines | Where-Object { $_ -notmatch '^Attribute VB_Name' })
            $body = ($filtered -join "`n")
            Invoke-Com { $cm.AddFromString($body) } "Inject $targetName"
            $lines.Add("injected document module code: " + $targetName)
        } else {
            $lines.Add("WARN: document component " + $targetName + " not found")
        }
    }

    # Compile by saving (Excel compiles the VBA project on save)
    Invoke-Com { $wb.Save() } "Save"
    $lines.Add("saved after import (Excel compiles on save)")

    # Export every component back for source-control verification
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

    # Hide the VBE before the final save so the saved candidate does not open
    # with the VBE window active (an open VBE blocks Application.Run under COM).
    try {
        Invoke-Com { $excel.VBE.MainWindow.Visible = $false } "VbeHide"
        $lines.Add("VBE window hidden before save")
    } catch { $lines.Add("VBE hide failed: " + $_.Exception.Message) }
    Invoke-Com { $wb.Save() } "SaveFinal"
    $lines.Add("final save after VBE hide")

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
