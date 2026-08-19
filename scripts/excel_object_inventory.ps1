# Generate workbook-object-inventory.txt via a fresh Excel COM session.
# Runs immediately after open (before any heavy operations) to avoid the
# degraded-COM state that plagues late-session enumeration.
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$WbPath = Join-Path $Root "workbook\LabInventory_v0.1.xlsx"
$OutDir = Join-Path $Root "evidence\excel-runtime"
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
function Invoke-Com([scriptblock]$block, [string]$label, [int]$attempts = 12) {
    for ($a = 1; $a -le $attempts; $a++) {
        try { return & $block } catch { if ($a -eq $attempts) { throw }; Start-Sleep -Milliseconds 700 }
    }
}
function Str($v) { if ($null -eq $v) { return "" }; try { return [string]$v } catch { return "?" } }
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = Invoke-Com { $excel.Workbooks.Open($WbPath, 0, $false) } "Open"
    $lines.Add("=== Workbook object inventory (Microsoft Excel COM) ===")
    $lines.Add("Workbook: $WbPath")
    $lines.Add("Excel: $($excel.Version) build $($excel.Build)")
    $lines.Add("")
    $lines.Add("WORKSHEETS:")
    $wsCount = [int](Invoke-Com { $wb.Worksheets.Count } "Count")
    for ($i = 1; $i -le $wsCount; $i++) {
        $ws = Invoke-Com { $wb.Worksheets.Item($i) } "Ws$i"
        $lines.Add("  " + (Str $ws.Name))
        $tables = @()
        $loCount = [int](Invoke-Com { $ws.ListObjects.Count } "Lo$i")
        for ($j = 1; $j -le $loCount; $j++) { $tables += (Invoke-Com { $ws.ListObjects.Item($j).Name } "LoN$i$j") }
        $lines.Add("    Tables: " + $(if ($tables.Count -gt 0) { $tables -join ', ' } else { 'none' }))
        $names = @()
        $nCount = [int](Invoke-Com { $ws.Names.Count } "Wn$i")
        for ($j = 1; $j -le $nCount; $j++) { $names += (Invoke-Com { $ws.Names.Item($j).Name } "WnN$i$j") }
        if ($names.Count -gt 0) { $lines.Add("    Names: $(($names -join ', '))") }
    }
    $lines.Add("")
    $lines.Add("WORKBOOK NAMED RANGES:")
    $nbCount = [int](Invoke-Com { $wb.Names.Count } "Nb")
    for ($i = 1; $i -le $nbCount; $i++) {
        $nm = Invoke-Com { $wb.Names.Item($i) } "Nb$i"
        $lines.Add("  " + (Str $nm.Name) + " -> " + (Str $nm.RefersTo))
    }
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines((Join-Path $OutDir "workbook-object-inventory.txt"), $lines)
    Write-Output "inventory written"
}
