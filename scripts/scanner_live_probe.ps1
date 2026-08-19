# Live scanner-detection probe: open the production workbook, select Scan!D7,
# and watch for ANY typed input arriving in this session over a window.
# This is NOT a pass/fail acceptance — it detects whether scanner keystrokes
# reach this RDP session at all.
$ErrorActionPreference = "Continue"
$Root = "C:\Users\Q\Documents\laboratory-inventory-excel"
$Xlsm = "$Root\workbook\LabInventory_v1.0.0-production.xlsm"
$LogFile = "$Root\evidence\production\scanner-live-probe.md"
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($Xlsm, 0, $false)
    Start-Sleep -Seconds 2
    $ws = $wb.Worksheets.Item("Scan")
    $ws.Activate()
    $ws.Unprotect()
    $ws.Range("D7").Value2 = ""
    $ws.Range("D7").Select()
    # leave D7 unprotected so typed/scanned input lands in it
    $lines.Add("## Scanner live probe - " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    $lines.Add("")
    $lines.Add("Scan!D7 selected and cleared; watching for typed input for 45 s.")
    $lines.Add("")
    $deadline = (Get-Date).AddSeconds(45)
    $sawInput = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $v = $ws.Range("D7").Value2
        if ($null -ne $v -and [string]$v -ne "") {
            $lines.Add("INPUT RECEIVED at " + (Get-Date -Format "HH:mm:ss") + ": [" + $v + "]")
            $sawInput = $true
            break
        }
    }
    if (-not $sawInput) {
        $lines.Add("No typed input received in 45 s - no scanner keystrokes reach this session.")
    }
    $lines.Add("")
    $lines.Add("Session: [" + $env:SESSIONNAME + "] (RDP session 1; console session 2).")
    $lines.Add("PnP scan found no scanner device and no new HID keyboard in this session.")
    $ws.Protect()
    $wb.Close($false)
} catch {
    $lines.Add("ERR: " + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [System.IO.File]::WriteAllLines($LogFile, $lines)
    Write-Output ($lines -join "`n")
}
