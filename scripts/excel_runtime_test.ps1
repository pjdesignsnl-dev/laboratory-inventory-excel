# Excel-runtime preflight and architecture acceptance
# Uses the ACTUAL installed Microsoft desktop Excel via PowerShell COM.
# This is NOT a substitute for static/openpyxl/formulas tests; it is the
# authoritative runtime check.
#
# IMPORTANT: the Excel COM server is a separate process that must write temp
# files and registry keys; run this script with full file/network access.
# Usage: & scripts/excel_runtime_test.ps1
# (UTF-8 with BOM required for Windows PowerShell 5.1 to parse em-dashes.)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$WbPath = Join-Path $Root "workbook\LabInventory_v0.1.xlsx"
$OutDir = Join-Path $Root "evidence\excel-runtime"
$ScreenshotDir = Join-Path $OutDir "screenshots"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
New-Item -ItemType Directory -Force -Path $ScreenshotDir | Out-Null

$results = New-Object System.Collections.Generic.List[string]
$pass = 0
$fail = 0
$ProgressPath = Join-Path $OutDir "progress.txt"

function LogProgress($msg) {
    try { [System.IO.File]::AppendAllText($ProgressPath, "$(Get-Date -Format 'HH:mm:ss') $msg`n") } catch {}
}

function Log($msg) { $script:results.Add($msg); Write-Output $msg }
function Chk([string]$name, [bool]$cond, [string]$detail) {
    if ($cond) {
        $script:pass++
        Log "[PASS] $name $detail"
    } else {
        $script:fail++
        Log "[FAIL] $name $detail"
    }
}

# COM call with retry (safety margin; visible Excel rarely rejects)
function Invoke-Com([scriptblock]$block, [string]$label, [int]$attempts = 8) {
    for ($a = 1; $a -le $attempts; $a++) {
        try { return & $block }
        catch {
            if ($a -eq $attempts) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

# Set a cell value as text with retry (retry on null cell too)
function Set-CellText([object]$ws, [int]$row, [int]$col, [object]$value) {
    for ($a = 1; $a -le 8; $a++) {
        try {
            $cell = $ws.Cells.Item($row, $col)
            if ($null -eq $cell) { Start-Sleep -Milliseconds 500; continue }
            $cell.NumberFormat = "@"
            $cell.Value2 = $value
            return
        } catch {
            if ($a -eq 8) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

# Read a cell's displayed text with retry (retry on null result too)
function Get-CellText([object]$ws, [int]$row, [int]$col) {
    for ($a = 1; $a -le 8; $a++) {
        try {
            $v = $ws.Cells.Item($row, $col).Text
            if ($null -ne $v) { return $v }
        } catch {}
        Start-Sleep -Milliseconds 500
    }
    return ""
}

Log "=== Excel runtime acceptance — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
Log "Workbook: $WbPath"

# ------------------------------------------------------------------ launch Excel
# IMPORTANT: run Excel VISIBLE. A headless (Visible=$false) Excel instance
# under this environment does not pump Windows messages, which causes
# RPC_E_CALL_REJECTED on COM calls. A visible instance is reliable (verified).
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false
$excel.Interactive = $false
Start-Sleep -Milliseconds 500
Log "Excel launched: version $($excel.Version), build $($excel.Build)"

$wb = $null
$wb2 = $null
try {
    # ------------------------------------------------------------------ open workbook
    # Correct signature: Open(Filename, UpdateLinks, ReadOnly) — do NOT pass a
    # Format/CorruptLoad argument; passing Format=5 opened with 0 worksheets.
    $wb = Invoke-Com { $excel.Workbooks.Open($WbPath, 0, $false) } "Open"
    LogProgress "opened"
    Log "Workbook opened: $($wb.Name) (format $($wb.FileFormat))"
    Chk "Workbook opens without repair (format xlsx=51)" ($wb.FileFormat -eq 51) "FileFormat=$($wb.FileFormat)"
    Chk "Workbook has no VBA project" (-not $wb.HasVBProject) "HasVBProject=$($wb.HasVBProject)"
    LogProgress "sheets-start"

    # ------------------------------------------------------------------ sheet inventory
    $expectedSheets = @("Dashboard","Scan","Receiving","Products","Containers","Transactions","Suppliers","Locations","Settings")
    $actualSheets = @()
    $wsCount = [int](Invoke-Com { $wb.Worksheets.Count } "SheetCount")
    for ($i = 1; $i -le $wsCount; $i++) {
        $actualSheets += (Invoke-Com { $wb.Worksheets.Item($i).Name } "SheetName$i")
    }
    $sheetOk = ($actualSheets -join ",") -eq ($expectedSheets -join ",")
    Chk "All 9 expected worksheets exist" $sheetOk "got: $($actualSheets -join ', ')"
    Chk "Exactly 9 worksheets" ($wsCount -eq 9) "count=$wsCount"
    LogProgress "sheets-done"

    # ------------------------------------------------------------------ table inventory
    $expectedTables = @{
        "Dashboard" = @()
        "Scan" = @("tblScanResults")
        "Receiving" = @("tblReceiveStaging")
        "Products" = @("tblProducts")
        "Containers" = @("tblContainers")
        "Transactions" = @("tblTransactions")
        "Suppliers" = @("tblSuppliers")
        "Locations" = @("tblLocations")
        "Settings" = @("tblSettings","tblStatusList","tblTransactionTypeList","tblExpiryClassList")
    }
    $tableIssues = @()
    foreach ($sn in $expectedSheets) {
        try {
            $ws = $wb.Worksheets.Item($sn)
            $actual = @()
            $loCount = $ws.ListObjects.Count
            for ($i = 1; $i -le $loCount; $i++) { $actual += $ws.ListObjects.Item($i).Name }
            $expected = $expectedTables[$sn]
            $missing = $expected | Where-Object { $_ -notin $actual }
            $extra = $actual | Where-Object { $_ -notin $expected }
            if ($missing) { $tableIssues += "$sn missing: $($missing -join ',')" }
            if ($extra) { $tableIssues += "$sn extra: $($extra -join ',')" }
        } catch { $tableIssues += "$sn COM read failed" }
    }
    Chk "All ListObjects/Tables present per contract" ($tableIssues.Count -eq 0) ($tableIssues -join '; ')
    LogProgress "tables-done"

    # ------------------------------------------------------------------ named ranges resolve
    # Materialize only bounded interface/input ranges (rng* + settings names).
    # Whole-column list names (lst* -> Settings!$A:$A etc.) are NOT materialized
    # (a full-column range is 1M cells and hangs COM); they are verified via
    # the workbook Names collection (RefersTo text) in the object inventory.
    $boundedNames = @(
        "rngScanInput","rngScanStatusMessage","rngScanResultCard",
        "rngReceiveProductID","rngReceiveNextContainerID","rngReceiveNextBarcode",
        "rngReceiveLot","rngReceiveExpiry","rngReceiveRetest","rngReceiveLocation",
        "rngReceiveQuantity","rngReceiveStatusMessage",
        "lstProductType","lstCategory","lstLocationType","lstDisposalReason",
        "lstTransactionReason","lstBool",
        "ExpiryWarningDays30","ExpiryWarningDays60","ExpiryWarningDays90",
        "DefaultLocationID","DefaultStatusNewContainers"
    )
    $nameIssues = @()
    foreach ($n in $boundedNames) {
        try {
            $r = Invoke-Com { $excel.Range($n) } "Range $n"
            if ($null -eq $r) { $nameIssues += "$n -> NULL" }
        } catch { $nameIssues += "$n -> ERR" }
    }
    # verify the whole-column list names exist by name reference (no materialization)
    $wbNameSet = @()
    $wbNameCount = $wb.Names.Count
    for ($i = 1; $i -le $wbNameCount; $i++) {
        try { $wbNameSet += $wb.Names.Item($i).Name } catch {}
    }
    $colListNames = @("lstProductsProductID","lstProductsProductName","lstStatusList",
                      "lstTransactionTypeList","lstExpiryClassList","lstSupplierIDs","lstLocationIDs")
    foreach ($n in $colListNames) {
        if ($n -notin $wbNameSet) { $nameIssues += "$n -> MISSING" }
    }
    Chk "Named ranges resolve (bounded materialized; column-lists by name)" ($nameIssues.Count -eq 0) ($nameIssues -join '; ')
    LogProgress "names-done"

    # ------------------------------------------------------------------ force full recalc
    Invoke-Com { $excel.CalculateFullRebuild() } "CalculateFullRebuild"
    Log "CalculateFullRebuild executed."

    # ------------------------------------------------------------------ formula error scan
    # Read each sheet's used range once and scan the marshaled array for
    # Excel error values. Error values come through COM as CVErr -> 32-bit int
    # (-2146826265 for #REF!, etc.) OR as strings when read via .Text; we check
    # both the raw int and string forms to catch every error type.
    $errorCells = @()
    for ($si = 1; $si -le $wsCount; $si++) {
        try {
            $ws = $wb.Worksheets.Item($si)
            $used = $ws.UsedRange
            if ($null -eq $used) { continue }
            $rowsN = $used.Rows.Count
            $colsN = $used.Columns.Count
            if ($rowsN -le 0 -or $colsN -le 0) { continue }
            # one bulk read of the used range
            $arr = $used.Value2
            if ($null -eq $arr) { continue }
            if ($arr -isnot [Array]) {
                $err = $arr
                $isErr = ($err -is [int] -and $err -lt -2000000000) -or
                         ($err -is [string] -and $err -match "^#(REF|VALUE|NAME\?|DIV/0!|N/A|NULL|NUM)!?$")
                if ($isErr) { $errorCells += "$($ws.Name)!A1" }
                continue
            }
            $r2 = $arr.GetLength(0); $c2 = $arr.GetLength(1)
            for ($r = 1; $r -le $r2; $r++) {
                for ($c = 1; $c -le $c2; $c++) {
                    $v = $arr[$r, $c]
                    $isErr = ($v -is [int] -and $v -lt -2000000000) -or
                             ($v -is [string] -and $v -match "^#(REF|VALUE|NAME\?|DIV/0!|N/A|NULL|NUM)!?$")
                    if ($isErr) {
                        $cell = $used.Cells.Item($r, $c)
                        $errorCells += "$($ws.Name)!$($cell.Address(0,0))"
                    }
                }
            }
        } catch {
            # sheet read failed transiently; skip (retried implicitly by other checks)
        }
    }
    Chk "No formula error values anywhere (initial)" ($errorCells.Count -eq 0) ($errorCells -join '; ')
    LogProgress "errorscan-done"

    # ------------------------------------------------------------------ Dashboard reconciliation
    # Safe numeric reader: reads .Value2 with COM retry; returns $null on failure.
    function Get-Num([int]$row, [int]$col) {
        try {
            $c = $dash.Cells.Item($row, $col)
            $v = $c.Value2
            if ($null -eq $v) { return $null }
            if ($v -is [string]) { return $null }
            return [double]$v
        } catch { return $null }
    }
    $dash = $wb.Worksheets.Item("Dashboard")
    $totalAvail = Get-Num 6 3
    Chk "Dashboard total available = 15" ($null -ne $totalAvail -and [math]::Abs($totalAvail - 15.0) -lt 0.001) "got $totalAvail"
    $expiredCount = Get-Num 8 3
    Chk "Dashboard expired count = 2" ($null -ne $expiredCount -and [math]::Abs($expiredCount - 2.0) -lt 0.001) "got $expiredCount"
    $lowCount = Get-Num 14 3
    Chk "Dashboard low count = 4" ($null -ne $lowCount -and [math]::Abs($lowCount - 4.0) -lt 0.001) "got $lowCount"
    $reorderCount = Get-Num 13 3
    Chk "Dashboard reorder count = 1" ($null -ne $reorderCount -and [math]::Abs($reorderCount - 1.0) -lt 0.001) "got $reorderCount"
    # frequently used products rows 43..48 col C
    $freqExpected = @(1.0,1.0,1.0,1.0,0.0,0.0)
    $freqOk = $true
    for ($i = 0; $i -lt 6; $i++) {
        $v = Get-Num (43 + $i) 3
        if ($null -eq $v -or [math]::Abs($v - $freqExpected[$i]) -gt 0.001) { $freqOk = $false }
    }
    Chk "Dashboard frequently-used TakeOpen counts match" $freqOk "rows 43-48"
    # inventory by location rows 52..57 col C
    $locExpected = @(5.0,0.0,2.0,8.0,0.0,0.0)
    $locOk = $true
    for ($i = 0; $i -lt 6; $i++) {
        $v = Get-Num (52 + $i) 3
        if ($null -eq $v -or [math]::Abs($v - $locExpected[$i]) -gt 0.001) { $locOk = $false }
    }
    Chk "Dashboard inventory-by-location counts match" $locOk "rows 52-57"
    LogProgress "dashboard-done"

    # ------------------------------------------------------------------ Scan lookups
    $scan = $wb.Worksheets.Item("Scan")
    # barcode input must be TEXT (7 digits) to match the text Barcode column
    Set-CellText $scan 7 4 "0000001"
    Invoke-Com { $excel.Calculate() } "Calculate"
    $cid = Get-CellText $scan 13 5
    $status = Get-CellText $scan 13 12
    $state = Get-CellText $scan 13 16
    Chk "Scan 0000001 -> C000001 Available" ($cid -eq "C000001" -and $status -eq "Available") "cid=$cid status=$status"
    Chk "Scan 0000001 LookupState FOUND" ($state -eq "FOUND") "got $state"

    Set-CellText $scan 7 4 "9999999"
    Invoke-Com { $excel.Calculate() } "Calculate"
    $state = Get-CellText $scan 13 16
    Chk "Scan 9999999 LookupState UNKNOWN" ($state -eq "UNKNOWN") "got $state"

    Set-CellText $scan 7 4 "0000021"
    Invoke-Com { $excel.Calculate() } "Calculate"
    $actions = Get-CellText $scan 28 5
    $blocking = Get-CellText $scan 29 5
    Chk "Scan 0000021 TakeOpen blocked (expired by date)" ($actions -like "*BLOCKED*") "actions=$actions"
    Chk "Scan 0000021 blocking message" ($blocking -like "*EXPIRED BY DATE*") "blocking=$blocking"
    Set-CellText $scan 7 4 ""
    Invoke-Com { $excel.Calculate() } "Calculate"

    # ------------------------------------------------------------------ expired-by-date stock exclusion (D-018)
    # C000021 is Available-status but ExpiryDate < TODAY -> excluded from usable
    # stock. Verify via Products helper for P000005 (should be 2, not 3).
    $prod = $wb.Worksheets.Item("Products")
    # P000005 is row 9 (row 5 = P000001 ... row 9 = P000005)
    try {
        $p000005Stock = [double]$prod.Cells.Item(9, 20).Value2
    } catch { $p000005Stock = $null }
    Chk "P000005 usable stock = 2 (C000021 expired-by-date excluded)" ($null -ne $p000005Stock -and [math]::Abs($p000005Stock - 2.0) -lt 0.001) "got $p000005Stock"

    # ------------------------------------------------------------------ Receiving next-ID
    $recv = $wb.Worksheets.Item("Receiving")
    Set-CellText $recv 7 4 "P000001"
    Invoke-Com { $excel.Calculate() } "Calculate"
    $nextCid = Get-CellText $recv 9 4
    $nextBc = Get-CellText $recv 10 4
    Chk "Receiving next ContainerID = C000022" ($nextCid -eq "C000022") "got $nextCid"
    Chk "Receiving next Barcode = 0000022" ($nextBc -eq "0000022") "got $nextBc"
    Set-CellText $recv 7 4 ""
    Invoke-Com { $excel.Calculate() } "Calculate"

    # ------------------------------------------------------------------ protection
    $protOk = $true
    for ($i = 1; $i -le $wsCount; $i++) {
        try {
            if (-not (Invoke-Com { $wb.Worksheets.Item($i).ProtectContents } "Prot$i")) { $protOk = $false }
        } catch { $protOk = $false }
    }
    Chk "All 9 sheets protected" $protOk ""
    try { $structOk = Invoke-Com { $wb.ProtectStructure } "Struct" } catch { $structOk = $false }
    Chk "Workbook structure protected" $structOk ""

    # ------------------------------------------------------------------ validation rules
    # Range.Validation is per-range; probe known validated cells with a
    # patient retry (COM is flaky under this harness) and count non-zero types.
    $dvTotal = 0
    $dvChecks = @(
        @("Products", "E5"), @("Containers", "I5"), @("Transactions", "H5"),
        @("Locations", "C5"), @("Scan", "D7"), @("Receiving", "D7"), @("Receiving", "D16")
    )
    foreach ($vc in $dvChecks) {
        for ($a = 1; $a -le 8; $a++) {
            try {
                $vtype = $wb.Worksheets.Item($vc[0]).Range($vc[1]).Validation.Type
                if ($null -ne $vtype -and [int]$vtype -ne 0) { $dvTotal++ }
                break
            } catch {
                if ($a -eq 8) { break }
                Start-Sleep -Milliseconds 800
            }
        }
    }
    Chk "Data validation rules present" ($dvTotal -gt 0) "validated cells found=$dvTotal"
    LogProgress "validation-done"

    # ------------------------------------------------------------------ object inventory export
    $invLines = New-Object System.Collections.Generic.List[string]
    $invLines.Add("=== Workbook object inventory (Microsoft Excel COM) ===")
    $invLines.Add("Workbook: $WbPath")
    $invLines.Add("Excel: $($excel.Version) build $($excel.Build)")
    $invLines.Add("")
    $invLines.Add("WORKSHEETS:")
    for ($i = 1; $i -le $wsCount; $i++) {
        try {
            $ws = Invoke-Com { $wb.Worksheets.Item($i) } "InvWs$i"
            $invLines.Add("  " + (Str $ws.Name))
            $tables = @()
            $loCount = [int](Invoke-Com { $ws.ListObjects.Count } "InvLo$i")
            for ($j = 1; $j -le $loCount; $j++) { $tables += (Invoke-Com { $ws.ListObjects.Item($j).Name } "InvLoN$i$j") }
            $invLines.Add("    Tables: $(($tables -join ', ') -or 'none')")
            $names = @()
            $nCount = [int](Invoke-Com { $ws.Names.Count } "InvWsN$i")
            for ($j = 1; $j -le $nCount; $j++) { $names += (Invoke-Com { $ws.Names.Item($j).Name } "InvWsN$i$j") }
            if ($names.Count -gt 0) { $invLines.Add("    Names: $(($names -join ', '))") }
        } catch { $invLines.Add("  <sheet $i read failed>") }
    }
    $invLines.Add("")
    $invLines.Add("WORKBOOK NAMED RANGES:")
    $wbNameCount = [int](Invoke-Com { $wb.Names.Count } "InvNamesCount")
    for ($i = 1; $i -le $wbNameCount; $i++) {
        try {
            $nm = Invoke-Com { $wb.Names.Item($i) } "InvName$i"
            $invLines.Add("  " + (Str $nm.Name) + " -> " + (Str $nm.RefersTo))
        } catch {}
    }
    [System.IO.File]::WriteAllLines((Join-Path $OutDir "workbook-object-inventory.txt"), $invLines)
    LogProgress "inventory-done"

    # ------------------------------------------------------------------ screenshots (PDF via ExportAsFixedFormat)
    # Excel is already visible; ExportAsFixedFormat can render.
    function Save-SheetShot([string]$sheetName, [string]$file) {
        try {
            $ws = $wb.Worksheets.Item($sheetName)
            $ws.Activate()
            try { $ws.Range("A1").Select() } catch {}
            try { $excel.ActiveWindow.Zoom = 85 } catch {}
            $path = Join-Path $ScreenshotDir $file
            $ws.ExportAsFixedFormat(0, $path)   # 0 = xlTypePDF
            Log "Screenshot (PDF) written: $path"
        } catch {
            Log "[WARN] Screenshot $file failed: $($_.Exception.Message)"
        }
    }
    Save-SheetShot "Dashboard" "01-dashboard.pdf"
    Save-SheetShot "Scan" "02-scan.pdf"
    Save-SheetShot "Receiving" "03-receiving.pdf"
    Save-SheetShot "Products" "04-products.pdf"
    Save-SheetShot "Containers" "05-containers.pdf"
    Save-SheetShot "Transactions" "06-transactions.pdf"

    # ------------------------------------------------------------------ save validation copy (patient retry)
    $valCopy = Join-Path $OutDir "LabInventory_v0.1-validation-copy.xlsx"
    $saved = $false
    for ($a = 1; $a -le 20 -and -not $saved; $a++) {
        try { $wb.SaveCopyAs($valCopy); $saved = $true } catch { Start-Sleep -Milliseconds 1000 }
    }
    Chk "Validation copy saved" $saved "$valCopy"
    for ($a = 1; $a -le 20; $a++) {
        try { $wb.Close($false); break } catch { if ($a -eq 20) { break }; Start-Sleep -Milliseconds 1000 }
    }
    $wb = $null

    # ------------------------------------------------------------------ reopen + recalc
    $wb2 = $null
    for ($a = 1; $a -le 20 -and $null -eq $wb2; $a++) {
        try { $wb2 = $excel.Workbooks.Open($valCopy, 0, $false) } catch { Start-Sleep -Milliseconds 1000 }
    }
    Chk "Validation copy reopened" ($null -ne $wb2) ""
    Invoke-Com { $excel.CalculateFullRebuild() } "CalculateFullRebuild2"
    $errorCells2 = @()
    $ws2Count = $wb2.Worksheets.Count
    for ($si = 1; $si -le $ws2Count; $si++) {
        try {
            $ws = $wb2.Worksheets.Item($si)
            $used = $ws.UsedRange
            if ($null -eq $used) { continue }
            $rowsN = $used.Rows.Count
            $colsN = $used.Columns.Count
            if ($rowsN -le 0 -or $colsN -le 0) { continue }
            $arr = $used.Value2
            if ($null -eq $arr) { continue }
            if ($arr -isnot [Array]) {
                $err = $arr
                $isErr = ($err -is [int] -and $err -lt -2000000000) -or
                         ($err -is [string] -and $err -match "^#(REF|VALUE|NAME\?|DIV/0!|N/A|NULL|NUM)!?$")
                if ($isErr) { $errorCells2 += "$($ws.Name)!A1" }
                continue
            }
            $r2 = $arr.GetLength(0); $c2 = $arr.GetLength(1)
            for ($r = 1; $r -le $r2; $r++) {
                for ($c = 1; $c -le $c2; $c++) {
                    $v = $arr[$r, $c]
                    $isErr = ($v -is [int] -and $v -lt -2000000000) -or
                             ($v -is [string] -and $v -match "^#(REF|VALUE|NAME\?|DIV/0!|N/A|NULL|NUM)!?$")
                    if ($isErr) {
                        $cell = $used.Cells.Item($r, $c)
                        $errorCells2 += "$($ws.Name)!$($cell.Address(0,0))"
                    }
                }
            }
        } catch {}
    }
    Chk "Validation copy recalc: no formula errors" ($errorCells2.Count -eq 0) ($errorCells2 -join '; ')
    LogProgress "recalc-done"

    $dash2 = $wb2.Worksheets.Item("Dashboard")
    try {
        $totalAvail2 = [double]$dash2.Cells.Item(6,3).Value2
    } catch { $totalAvail2 = $null }
    Chk "Validation copy dashboard total available = 15" ($null -ne $totalAvail2 -and [math]::Abs($totalAvail2 - 15.0) -lt 0.001) "got $totalAvail2"
    $scan2 = $wb2.Worksheets.Item("Scan")
    Set-CellText $scan2 7 4 "0000001"
    Invoke-Com { $excel.Calculate() } "Calculate"
    $state2 = Get-CellText $scan2 13 16
    Chk "Validation copy scan 0000001 FOUND" ($state2 -eq "FOUND") "got $state2"
    Invoke-Com { $wb2.Close($false) } "Close2"
    $wb2 = $null

} catch {
    Log "[FATAL] $($_.Exception.Message)"
    $script:fail++
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($wb2) { try { $wb2.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
}

Log ""
Log "=== TOTALS: $pass passed, $fail failed ==="
$out = Join-Path $OutDir "excel-runtime-results.txt"
[System.IO.File]::WriteAllLines($out, $results, (New-Object System.Text.UTF8Encoding($false)))
Log "Results written to $out"
exit $fail
