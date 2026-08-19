# Save/close/reopen verification: open candidate, modify (receive one), save,
# close, reopen, verify the mutation persisted and contract still validates.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0-candidate.xlsm"
$LogFile = "$Root\evidence\vba\save-reopen-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $LogFile -ErrorAction SilentlyContinue
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    $lines.Add("opened")
    $before = $wb.Worksheets.Item("Containers").ListObjects.Item("tblContainers").DataBodyRange.Rows.Count
    $lines.Add("containers-before=" + $before)
    # receive one container via the engine
    $excel.Run("modTestHooks.Test_ReceiveOne")
    $after = $wb.Worksheets.Item("Containers").ListObjects.Item("tblContainers").DataBodyRange.Rows.Count
    $lines.Add("containers-after=" + $after)
    # save + close
    $wb.Save()
    $wb.Close($false)
    $lines.Add("saved and closed")
    # reopen
    $wb2 = $excel.Workbooks.Open($Xlsm, 0, $false)
    $reopened = $wb2.Worksheets.Item("Containers").ListObjects.Item("tblContainers").DataBodyRange.Rows.Count
    $lines.Add("containers-reopened=" + $reopened)
    $lines.Add("persisted=" + ($reopened -eq $after))
    # contract still validates on the reopened file
    $excel.Run("modTestHooks.Test_ContractCheck")
    $wb2.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
