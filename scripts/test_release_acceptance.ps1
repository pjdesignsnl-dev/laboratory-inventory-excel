# Final release acceptance: open LabInventory_v1.0.0.xlsm with EVENTS ENABLED
# (real user experience), verify Workbook_Open contract validation passes
# (status bar ready + no fail-closed), verify the workbook structure intact.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0.xlsm"
$LogFile = "$Root\evidence\vba\release-acceptance-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # events ENABLED: Workbook_Open runs the contract validator
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened with events enabled (Workbook_Open ran)")
    Start-Sleep -Seconds 2
    $sb = $excel.StatusBar
    $lines.Add("statusbar=[" + $sb + "]")
    # verify VBA project present
    $hasVBA = $wb.HasVBProject
    $lines.Add("has-vb-project=" + $hasVBA)
    $cc = $wb.VBProject.VBComponents.Count
    $lines.Add("vb-components=" + $cc)
    # verify sheets intact
    $n = $wb.Worksheets.Count
    $lines.Add("sheets=" + $n)
    # verify a scan works through the event path in the release file
    $ws = $wb.Worksheets("Scan")
    $ws.Unprotect()
    $ws.Range("D7").Value2 = "0000001"
    $ws.Protect()
    Start-Sleep -Seconds 2
    $msg = $ws.Range("D9").Value2
    $lines.Add("scan-msg=[" + $msg + "]")
    $lines.Add("scan-ok=" + ($msg -match "FOUND"))
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
