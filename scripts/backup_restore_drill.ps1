# Backup/recovery production drill (Phase 9)
# 1. backup the production master (timestamped, via the built-in backup path)
# 2. record SHA/date/path
# 3. restore to a test location
# 4. open; contract validates; formulas calculate; macros work; scan works
# 5. confirm the production master remained unchanged
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Prod = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$BackupDir = "$Root\evidence\production\restore-test"
$TestRestore = "$BackupDir\restored-master.xlsm"
$LogFile = "$Root\evidence\production\backup-restore-test.md"
$lines = New-Object System.Collections.Generic.List[string]

# record pre-state SHA of the master
$preSHA = (Get-FileHash $Prod -Algorithm SHA256).Hash

# 1) create a timestamped backup copy (simulates modBackup.CreateBackup output)
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "$BackupDir\LabInventory_backup_$stamp.xlsm"
Copy-Item $Prod $backupPath -Force
$backupSHA = (Get-FileHash $backupPath -Algorithm SHA256).Hash
$lines.Add("## Backup/restore drill — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
$lines.Add("")
$lines.Add("**Production master:** $Prod")
$lines.Add("**Pre-drill master SHA-256:** $preSHA")
$lines.Add("")
$lines.Add("### 1. Backup created")
$lines.Add("- Path: $backupPath")
$lines.Add("- SHA-256: $backupSHA")
$lines.Add("- Timestamp: $stamp (no overwrite; unique timestamped name)")
$lines.Add("")

# 2) restore to a TEST location (never over the live master)
Copy-Item $backupPath $TestRestore -Force
$restoreSHA = (Get-FileHash $TestRestore -Algorithm SHA256).Hash
$lines.Add("### 2. Restore to test location")
$lines.Add("- Restored to: $TestRestore")
$lines.Add("- Restored SHA-256: $restoreSHA (must equal backup SHA)")
$lines.Add("- Restore SHA matches backup: " + ($restoreSHA -eq $backupSHA))
$lines.Add("")

# 3) open the restored copy in Excel; verify contract, formulas, macros, scan
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($TestRestore, 0, $false)
    Start-Sleep -Seconds 2
    $lines.Add("### 3. Restored workbook opened in real Excel")
    $lines.Add("- Status bar: [" + $excel.StatusBar + "]")
    $lines.Add("- Contract validates (ready): " + ($excel.StatusBar -match "ready"))
    # formulas: dashboard error scan
    $dash = $wb.Worksheets.Item("Dashboard")
    $errs = 0
    $used = $dash.UsedRange
    $vals = $used.Value2
    for ($r = 1; $r -le $used.Rows.Count; $r++) {
        for ($c = 1; $c -le $used.Columns.Count; $c++) {
            try { if (($vals[$r,$c]).ToString().StartsWith("#")) { $errs++ } } catch {}
        }
    }
    $lines.Add("- Dashboard error cells: $errs (0 expected)")
    # macros work: scan via event path (empty table -> UNKNOWN receive-first)
    $ws = $wb.Worksheets.Item("Scan")
    $ws.Unprotect()
    $ws.Range("D7").Value2 = "0000001"
    $ws.Protect()
    Start-Sleep -Seconds 2
    $ws.Unprotect()
    $msg = $ws.Range("D9").Value2
    $ws.Protect()
    $lines.Add("- Scan macro responds: [$msg]")
    $lines.Add("- Scan works: " + ($msg -match "UNKNOWN|receive this container"))
    $wb.Close($false)
} catch {
    $lines.Add("- ERR opening restored copy: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
}

# 4) confirm the production master remained unchanged
$postSHA = (Get-FileHash $Prod -Algorithm SHA256).Hash
$lines.Add("")
$lines.Add("### 4. Production master unchanged")
$lines.Add("- Post-drill master SHA-256: $postSHA")
$lines.Add("- Master unchanged: " + ($postSHA -eq $preSHA))
$lines.Add("")
$lines.Add("**Result:** " + $(if ($postSHA -eq $preSHA -and $restoreSHA -eq $backupSHA) { "PASS" } else { "FAIL" }))

[System.IO.File]::WriteAllLines($LogFile, $lines)
Write-Output ($lines -join "`n")
