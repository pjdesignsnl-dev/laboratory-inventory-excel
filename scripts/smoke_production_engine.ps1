# Production engine smoke via a throwaway driver module imported into a
# production-binary COPY (never the committed binary). Requires AccessVBOM=1
# for the import; restored to 0 immediately after.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Src = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$Copy = "$Root\workbook\LabInventory_v1.0.0-production-smokecopy.xlsm"
$Driver = "$Root\.tools\eval\modProdDriver.bas"
$OutFile = "$Root\evidence\vba\prod-drive-out.txt"
$LogFile = "$Root\evidence\vba\production-engine-smoke-log.txt"
$lines = New-Object System.Collections.Generic.List[string]
Remove-Item $OutFile -ErrorAction SilentlyContinue
Remove-Item $Copy -ErrorAction SilentlyContinue
Copy-Item $Src $Copy -Force
$lines.Add("copied production binary to smoke copy")

# temporarily enable AccessVBOM for the module import
$key = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($Copy, 0, $false)
    $lines.Add("opened smoke copy")
    $wb.VBProject.VBComponents.Import($Driver)
    $wb.Save()
    $lines.Add("imported throwaway driver")

    $excel.Run("modProdDriver.Drive_ReceiveOne")
    $lines.Add("Drive_ReceiveOne done")
    $excel.Run("modProdDriver.Drive_TakeOpenReturn")
    $lines.Add("Drive_TakeOpenReturn done")

    if (Test-Path $OutFile) { foreach ($l in Get-Content $OutFile) { $lines.Add("OUT: " + $l) } }
    else { $lines.Add("driver out file NOT created") }

    $wb.Save()
    $wb.Close($false)
    $lines.Add("saved and closed smoke copy")

    # reopen to confirm persistence
    $wb2 = $excel.Workbooks.Open($Copy, 0, $false)
    $cc = $wb2.Worksheets.Item("Containers").ListObjects.Item("tblContainers").DataBodyRange.Rows.Count
    $lines.Add("reopen-containers=" + $cc)
    $wb2.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    Set-ItemProperty -Path $key -Name AccessVBOM -Value 0 -Type DWord
    $lines.Add("AccessVBOM restored to 0")
    Remove-Item $Copy -ErrorAction SilentlyContinue
    Remove-Item $OutFile -ErrorAction SilentlyContinue
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
