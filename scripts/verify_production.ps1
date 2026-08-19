# Production workbook verification (AccessVBOM=0): open, Workbook_Open
# contract validation, verify tables are empty of fixtures, verify controlled
# lists preserved, verify scan workflow responds.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$LogFile = "$Root\evidence\vba\production-verify-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened with events enabled")
    Start-Sleep -Seconds 2
    $lines.Add("statusbar=[" + $excel.StatusBar + "]")

    foreach ($t in @(@("Products","tblProducts"), @("Containers","tblContainers"), @("Transactions","tblTransactions"), @("Suppliers","tblSuppliers"), @("Locations","tblLocations"))) {
        $ws = $wb.Worksheets.Item($t[0])
        $lo = $ws.ListObjects.Item($t[1])
        if ($lo.DataBodyRange) { $cnt = $lo.DataBodyRange.Rows.Count } else { $cnt = 0 }
        $lines.Add($t[1] + "-rows=" + $cnt)
    }
    $st = $wb.Worksheets.Item("Settings")
    $sl = $st.ListObjects.Item("tblStatusList")
    if ($sl.DataBodyRange) { $sc = $sl.DataBodyRange.Rows.Count } else { $sc = 0 }
    $lines.Add("tblStatusList-rows=" + $sc)
    $tl = $st.ListObjects.Item("tblTransactionTypeList")
    if ($tl.DataBodyRange) { $tc = $tl.DataBodyRange.Rows.Count } else { $tc = 0 }
    $lines.Add("tblTransactionTypeList-rows=" + $tc)

    # scan workflow responds (barcode 0000001 with empty container table -> UNKNOWN)
    $ws = $wb.Worksheets.Item("Scan")
    $ws.Unprotect()
    $ws.Range("D7").Value2 = "0000001"
    $ws.Protect()
    Start-Sleep -Seconds 2
    $ws.Unprotect()
    $msg = $ws.Range("D9").Value2
    $ws.Protect()
    $lines.Add("scan-unknown-msg=[" + $msg + "]")
    $lines.Add("scan-responds=" + ($msg -match "UNKNOWN|receive this container"))

    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
