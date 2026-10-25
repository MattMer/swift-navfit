Attribute VB_Name = "BackupHarness"
Option Explicit
Public Function PrepareBackupCheckpoint() As String
    Dim id As Long, block As Long, entry As Long
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC BACKUP", "Synthetic backup tester")
    id = CommitDraft(0, "CHIEFEVAL", "Saved backup title", "Synthetic Member", Null, Null, "", "Initial")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtTitle = "Pending backup title"
    Forms!frmEditor!txtStart = "not-a-date"
    For block = 33 To 39
        For entry = 1 To 3
            Forms!frmEditor.Controls("txtTrait" & block & "Comment" & entry) = "BACKUP COMMENT " & block & "." & entry
        Next
    Next
    If Not CheckpointEditor() Then Err.Raise 5, , "Pending checkpoint failed"
    If DCount("*", "RecoveryDraft") <> 1 Then Err.Raise 5, , "Pending checkpoint missing"
    If MajorScalar("SELECT Title FROM Report WHERE ReportID=" & id) <> "Saved backup title" Then Err.Raise 5, , "Checkpoint changed committed report"
    PrepareBackupCheckpoint = "PASS: checkpoint includes invalid raw date without committing it": Exit Function
Failed:
    PrepareBackupCheckpoint = "FAIL: " & Err.Description
End Function
'Catch intentional rejections so Access shows no run-time dialog to the COM runner.
Public Function TryWorkspaceBackup(ByVal destination As String) As String
    On Error GoTo Failed
    TryWorkspaceBackup = WriteWorkspaceBackup(destination): Exit Function
Failed:
    TryWorkspaceBackup = "REJECTED: " & Err.Description
End Function
