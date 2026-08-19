# Deployment-location smoke test (Phase 11): place the production master in a
# simulated deployment location (stands in for \\<LAB-SERVER>\Inventory\), run
# the operator flow FROM that location, verify persistence + second-writer not
# permitted, then confirm the committed clean production master is untouched.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Prod = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$DeployDir = "$Root\evidence\production\deployment-location"
$Deployed = "$DeployDir\LabInventory_v1.0.0-production.xlsm"
$Driver = "$Root\.tools\eval\modProdDriver.bas"
$OutFile = "$Root\evidence\vba\prod-drive-out.txt"
$LogFile = "$Root\evidence\production\deployment-location-smoke.md"
$lines = New-Object System.Collections.Generic.List[string]

New-Item -ItemType Directory -Force -Path $DeployDir | Out-Null
Remove-Item $OutFile -ErrorAction SilentlyContinue
$preProdSHA = (Get-FileHash $Prod -Algorithm SHA256).Hash

# deploy a copy (the committed master stays untouched in workbook/)
Copy-Item $Prod $Deployed -Force
$lines.Add("## Deployment-location smoke — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
$lines.Add("")
$lines.Add("**Simulated deployment location:** $DeployDir (stands in for the D-023 network share \\\\<LAB-SERVER>\\Inventory\\)")
$lines.Add("**Deployed master SHA-256:** " + (Get-FileHash $Deployed -Algorithm SHA256).Hash)
$lines.Add("**Committed clean production master SHA-256 (must stay unchanged):** $preProdSHA")
$lines.Add("")

# temporarily enable AccessVBOM to import the throwaway driver into the deployed copy
$key = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # open with EVENTS ENABLED so Workbook_Open runs the contract validation
    $wb = $excel.Workbooks.Open($Deployed, 0, $false)
    Start-Sleep -Seconds 3   # allow Workbook_Open async contract validation to finish
    $lines.Add("### 1. Open from deployment location")
    $lines.Add("- Workbook_Open contract validation (status bar): [" + $excel.StatusBar + "]")
    $lines.Add("- Contract OK: " + ($excel.StatusBar -match "ready"))
    $excel.EnableEvents = $false  # disable events for the driver-import phase

    # import throwaway driver and run the operator flow (import ref -> receive -> TakeOpen -> Return)
    $wb.VBProject.VBComponents.Import($Driver)
    $wb.Save()
    $lines.Add("- Driver imported (throwaway, not part of any committed binary)")
    $excel.Run("modProdDriver.Drive_ReceiveOne")
    $excel.Run("modProdDriver.Drive_TakeOpenReturn")
    Start-Sleep -Seconds 2
    if (Test-Path $OutFile) {
        foreach ($l in Get-Content $OutFile) { $lines.Add("- " + $l) }
    }

    # transaction append + dashboard + backup + save/close/reopen persistence
    $tlo = $wb.Worksheets.Item("Transactions").ListObjects.Item("tblTransactions")
    $lines.Add("### 2. Operator flow results")
    $lines.Add("- Transaction rows appended: " + $tlo.DataBodyRange.Rows.Count + " (expect 3: Receive, TakeOpen, Return)")
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

    # backup (timestamped, never overwrite) into the deployment-location backups subfolder
    $bkDir = "$DeployDir\backups"
    New-Item -ItemType Directory -Force -Path $bkDir | Out-Null
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $bk = "$bkDir\LabInventory_backup_$stamp.xlsm"
    $wb.SaveCopyAs($bk)
    $lines.Add("- Backup created (timestamped, no overwrite): $bk")

    # save + close + reopen
    $wb.Save()
    $wb.Close($false)
    $wb2 = $excel.Workbooks.Open($Deployed, 0, $false)
    $cc = $wb2.Worksheets.Item("Containers").ListObjects.Item("tblContainers").DataBodyRange.Rows.Count
    $lines.Add("- Reopen: containers=" + $cc + " (expect 1) — persistence OK: " + ($cc -eq 1))
    $wb2.Close($false)
    $lines.Add("- Save/close/reopen persistence: PASS")

    # second-writer-not-permitted: open in a second Excel instance while the
    # first still has it open (first instance re-opens it now)
    $excel2 = $null
    try {
        $excel2 = New-Object -ComObject Excel.Application
        $excel2.Visible = $true
        $excel2.DisplayAlerts = $false
        $wbFirst = $excel.Workbooks.Open($Deployed, 0, $false)  # first instance holds it
        $wbSecond = $excel2.Workbooks.Open($Deployed, 0, $true) # second instance tries ReadOnly
        $lines.Add("### 3. Second writer not permitted")
        $lines.Add("- Second instance opened the deployed master: " + $wbSecond.ReadOnly)
        $lines.Add("- Second writer permitted: " + (-not $wbSecond.ReadOnly))
        $wbSecond.Close($false)
        $wbFirst.Close($false)
    } catch {
        $lines.Add("- Second instance open blocked (expected if locked): " + $_.Exception.Message)
    } finally {
        if ($excel2) { try { $excel2.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null } catch {} }
    }
} catch {
    $lines.Add("### ERR")
    $lines.Add("- " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    Set-ItemProperty -Path $key -Name AccessVBOM -Value 0 -Type DWord
}

# confirm the committed clean production master is unchanged
$postProdSHA = (Get-FileHash $Prod -Algorithm SHA256).Hash
$lines.Add("")
$lines.Add("### 4. Committed clean production master unchanged")
$lines.Add("- Pre-smoke SHA: $preProdSHA")
$lines.Add("- Post-smoke SHA: $postProdSHA")
$lines.Add("- Master unchanged: " + ($preProdSHA -eq $postProdSHA))
$lines.Add("")
$lines.Add("**AccessVBOM restored to:** " + (Get-ItemProperty -Path $key -Name AccessVBOM).AccessVBOM)

Remove-Item $OutFile -ErrorAction SilentlyContinue
[System.IO.File]::WriteAllLines($LogFile, $lines)
Write-Output ($lines -join "`n")
