Attribute VB_Name = "RecoveryHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 950, , message
    checks = checks + 1
End Sub
Private Function RejectStale(ByVal recoveryID As Long) As Boolean
    On Error GoTo Rejected
    Call RestoreRecoveryDraft(recoveryID)
    Exit Function
Rejected:
    RejectStale = InStr(Err.Description, "changed") > 0
End Function
Private Function RejectGuard() As Boolean
    On Error GoTo Rejected
    Call GuardEditorBaseline()
    Exit Function
Rejected:
    RejectGuard = InStr(Err.Description, "newer revision") > 0
End Function
Public Function PrepareRecovery() As String
    Dim id As Long, r As DAO.Recordset, f As Form, text As String
    On Error GoTo Failed
    checks = 0: Call ConfigureWorkspace("SYNTHETIC RECOVERY", "Synthetic Tester")
    id = CommitDraft(0, "EVAL", "Saved title", "Synthetic", Null, Null, "Saved notes", "Initial", "Saved narrative")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Set f = Forms!frmEditor
    Check f.TimerInterval = 30000 And f.OnTimer = "=CheckpointEditor()", "Automatic checkpoint not wired"
    Check CheckpointEditor(), "Initial checkpoint check failed"
    Check DCount("*", "RecoveryDraft") = 0, "Untouched editor created recovery"
    text = "Pending narrative" & vbCrLf & String$(7000, "x")
    f!txtTitle = "Pending title": f!txtOfficialComments = text: f!txtStart = "not-a-date"
    Check CheckpointEditor(), "Checkpoint failed"
    Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Checkpoint committed revision"
    Check DLookup("Title", "Report", "ReportID=" & id) = "Saved title", "Checkpoint changed report"
    f!txtTitle = "Pending title v2": Check CheckpointEditor(), "Second checkpoint failed"
    Check DCount("*", "RecoveryDraft") = 1, "Checkpoint duplicated session"
    CurrentDb.TableDefs("RecoveryDraft").Fields("DraftTitle").ValidationRule = "<> 'Rejected checkpoint'"
    f!txtTitle = "Rejected checkpoint"
    Check Not CheckpointEditor(), "Injected checkpoint failure accepted"
    Check DLookup("DraftTitle", "RecoveryDraft") = "Pending title v2", "Failed checkpoint destroyed previous edits"
    CurrentDb.TableDefs("RecoveryDraft").Fields("DraftTitle").ValidationRule = ""
    f!txtTitle = "Pending title v2"
    'Closing without the explicit discard action represents an interrupted editor session.
    DoCmd.Close acForm, "frmEditor", acSaveNo
    DoCmd.OpenForm "frmEditor", , , , , , "0": Set f = Forms!frmEditor
    f!cboKind = "CHIEFEVAL": f!txtTitle = "Unsaved new draft": f!txtNotes = "Uncommitted notes"
    Check CheckpointEditor(), "New draft checkpoint failed"
    Check DCount("*", "Report") = 1 And DCount("*", "RecoveryDraft") = 2, "New checkpoint created report"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    PrepareRecovery = "PASS: " & checks & " preparation assertions; two checkpoints retained"
    Exit Function
Failed:
    PrepareRecovery = "FAIL: " & Err.Number & " " & Err.Description
End Function

Public Function PrepareTimerCheck() As Boolean
    Dim id As Long
    id = DLookup("ReportID", "Report")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor.SetFocus
    Forms!frmEditor!txtTitle.SetFocus
    Forms!frmEditor!txtTitle.Text = "Timer captured active text"
    Forms!frmEditor.TimerInterval = 100
    PrepareTimerCheck = True
End Function
Public Function FinishTimerCheck() As Boolean
    Dim id As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If Nz(DLookup("DraftTitle", "RecoveryDraft", "ReportID=" & id), "") <> "Timer captured active text" Then Err.Raise vbObjectError + 951, , "Timer did not capture active textbox text."
    Forms!frmEditor.TimerInterval = 30000
    Call DiscardEditorCheckpoint()
    DoCmd.Close acForm, "frmEditor", acSaveNo
    FinishTimerCheck = True
End Function
Public Function VerifyRecovery() As String
    Dim recoveryID As Long, id As Long, f As Form, text As String, other As Long
    On Error GoTo Failed
    checks = 0
    Check DCount("*", "RecoveryDraft") = 2, "Checkpoints lost on reopen"
    id = DLookup("ReportID", "RecoveryDraft", "ReportID Is Not Null")
    recoveryID = DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)
    Check RestoreRecoveryDraft(recoveryID), "Existing recovery failed"
    Set f = Forms!frmEditor
    Check Screen.ActiveForm.Name = "frmEditor", "Recovered editor not active"
    Check f!txtTitle = "Pending title v2", "Latest title lost"
    Check f!txtStart = "not-a-date", "Incomplete date lost"
    text = "Pending narrative" & vbCrLf & String$(7000, "x")
    Check StrComp(f!txtOfficialComments, text, vbBinaryCompare) = 0, "Long narrative lost"
    Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Recovery committed values"
    f!txtStart = Null
    Check SaveEditor(), "Recovered save failed"
    Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Recovery save missing revision"
    Check DCount("*", "RecoveryDraft") = 1, "Saved checkpoint not cleared"
    Check CheckpointEditor() And DCount("*", "RecoveryDraft") = 1, "Saved editor recreated checkpoint"
    f!txtTitle = "Stale pending title": Check CheckpointEditor(), "Stale checkpoint failed"
    recoveryID = DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)
    Call CommitDraft(id, "EVAL", "Newer committed title", "Synthetic", Null, Null, "Saved notes", "External edit", "Newer narrative")
    Check RejectGuard(), "Editor could overwrite newer revision"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Check RejectStale(recoveryID), "Stale recovery accepted"
    Check RestoreRecoveryDraft(recoveryID, True), "Stale recovery as new failed"
    Check Forms!frmEditor!txtReportID = 0, "Recovery as new reused identity"
    Check SaveEditor(), "Save as new failed"
    Check DLookup("Title", "Report", "ReportID=" & id) = "Newer committed title", "Newer revision overwritten"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    recoveryID = DLookup("RecoveryDraftID", "RecoveryDraft")
    Check RestoreRecoveryDraft(recoveryID), "New unsaved draft recovery failed"
    Check Forms!frmEditor!cboKind = "CHIEFEVAL" And Forms!frmEditor!txtTitle = "Unsaved new draft", "New draft values lost"
    Check DiscardEditorCheckpoint(), "Explicit discard failed"
    Check DCount("*", "RecoveryDraft") = 0, "Discard retained checkpoint"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    VerifyRecovery = "PASS: " & checks & " recovery assertions"
    Exit Function
Failed:
    VerifyRecovery = "FAIL: " & Err.Number & " " & Err.Description
End Function
