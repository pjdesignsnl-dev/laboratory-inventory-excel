# Read-only report acceptance (Phase 10): open the report, verify no VBA,
# banner present, protected sheets, viewer content (dashboard/stock/expiry/
# locations) readable, and no write path.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0-readonly-report.xlsx"
$LogFile = "$Root\evidence\vba\readonly-acceptance-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened read-only report")
    $lines.Add("has-vba-project=" + $wb.HasVBProject)  # must be False
    $lines.Add("sheets=" + $wb.Worksheets.Count)

    # banner
    $dash = $wb.Worksheets.Item("Dashboard")
    $dash.Unprotect()
    $banner = $dash.Cells.Item(1, 1).Value2
    $dash.Protect()
    $lines.Add("banner=[" + $banner + "]")

    # protected?
    $lines.Add("dashboard-protected=" + $dash.ProtectContents)
    $scan = $wb.Worksheets.Item("Scan")
    $lines.Add("scan-protected=" + $scan.ProtectContents)

    # viewer content present: Dashboard labels + tables
    $errs = 0
    $used = $dash.UsedRange
    $vals = $used.Value2
    for ($r = 1; $r -le $used.Rows.Count; $r++) {
        for ($c = 1; $c -le $used.Columns.Count; $c++) {
            try { if (($vals[$r,$c]).ToString().StartsWith("#")) { $errs++ } } catch {}
        }
    }
    $lines.Add("dashboard-error-cells=" + $errs)
    # expiry/reorder/location tables present
    foreach ($t in @(@("Containers","tblContainers"), @("Locations","tblLocations"))) {
        $ws = $wb.Worksheets.Item($t[0])
        $lo = $ws.ListObjects.Item($t[1])
        $cnt = if ($lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
        $lines.Add($t[1] + "-rows=" + $cnt)
    }
    # no write path: workbook structure protected (cannot add sheets)
    $lines.Add("structure-protected=" + $wb.ProtectStructure)

    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
