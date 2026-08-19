Attribute VB_Name = "modBackup"
Option Explicit

' ============================================================================
' modBackup - Backup / recovery (Stage 11)
' ============================================================================
' Timestamped unique backups. A required backup that fails is treated as a
' failed operation (no silent continue). Never overwrite the only known-good
' backup.
' ============================================================================

Private Const BACKUP_FOLDER_ENV As String = "LABINV_BACKUP_FOLDER"

Public Function BackupFolder() As String
    Dim f As String
    f = Environ$(BACKUP_FOLDER_ENV)
    If Len(f) = 0 Then
        f = ThisWorkbook.Path
    End If
    BackupFolder = f
End Function

Public Function CreateBackup(Optional ByVal note As String = "") As String
    ' Creates a timestamped copy of the current workbook in the backup folder.
    ' Returns the backup path. Raises on failure (caller treats as failure).
    Dim folder As String
    folder = BackupFolder()
    If Len(folder) = 0 Then
        Err.Raise vbObjectError + 2601, "modBackup", "No backup folder configured."
    End If

    Dim stamp As String
    stamp = Format$(modUtilities.GetNow(), "yyyymmdd_hhnnss")
    Dim baseName As String
    baseName = "LabInventory_backup_" & stamp
    Dim path As String
    path = folder & Application.PathSeparator & baseName & ".xlsm"

    ' never overwrite an existing backup
    Dim n As Long
    n = 1
    Do While Len(Dir$(path)) > 0
        n = n + 1
        path = folder & Application.PathSeparator & baseName & "_" & n & ".xlsm"
    Loop

    On Error GoTo failBackup
    Application.EnableEvents = False
    ThisWorkbook.SaveCopyAs path
    Application.EnableEvents = True
    CreateBackup = path
    Exit Function

failBackup:
    Application.EnableEvents = True
    Err.Raise vbObjectError + 2602, "modBackup", _
              "Backup failed and the operation was not continued: " & Err.Description
End Function

Public Sub RequireBackupBeforeMutation()
    ' Call before any mutation that requires a successful backup first.
    Dim path As String
    path = CreateBackup("pre-mutation backup")
    ' If CreateBackup raised, this sub propagates the error, failing the
    ' intended operation (no silent continue).
End Sub

Public Function RestoreFromBackup(ByVal backupPath As String) As Boolean
    ' Recovery instructions are operator-driven (cannot replace the open
    ' workbook safely while it is running). This validates the backup file.
    If Len(Dir$(backupPath)) = 0 Then
        RestoreFromBackup = False
        Exit Function
    End If
    RestoreFromBackup = True
End Function

Public Function LatestBackup() As String
    ' Returns the most recent backup file in the folder (by name timestamp).
    Dim folder As String
    folder = BackupFolder()
    Dim f As String
    f = Dir$(folder & Application.PathSeparator & "LabInventory_backup_*.xlsm")
    Dim best As String
    best = ""
    Do While Len(f) > 0
        If StrComp(f, best, vbTextCompare) > 0 Then best = f
        f = Dir$
    Loop
    If Len(best) > 0 Then LatestBackup = folder & Application.PathSeparator & best
End Function
