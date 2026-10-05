Attribute VB_Name = "BackupHarness"
Option Explicit
Public Function PrepareBackupCheckpoint() As String
    Dim id As Long
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC BACKUP", "Synthetic backup tester")
    id = CommitDraft(0, "EVAL", "Saved backup title", "Synthetic Member", Null, Null, "", "Initial")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtTitle = "Pending backup title"
    Forms!frmEditor!txtStart = "not-a-date"
    If Not CheckpointEditor() Then Err.Raise 5, , "Pending checkpoint failed"
    If DCount("*", "RecoveryDraft") <> 1 Then Err.Raise 5, , "Pending checkpoint missing"
    If MajorScalar("SELECT Title FROM Report WHERE ReportID=" & id) <> "Saved backup title" Then Err.Raise 5, , "Checkpoint changed committed report"
    PrepareBackupCheckpoint = "PASS: checkpoint includes invalid raw date without committing it": Exit Function
Failed:
    PrepareBackupCheckpoint = "FAIL: " & Err.Description
End Function
