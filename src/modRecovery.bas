Attribute VB_Name = "modRecovery"
Option Compare Database
Option Explicit
Private sessionKey As String
Private baseRevision As Long
Private checkpointSignature As String
Private baseline(0 To 78) As Variant
' Keep the editor control order aligned with DraftFieldKeys and recovery storage fields.
Private Function ControlsList() As Variant
    ControlsList = Array("cboKind", "txtTitle", "txtMember", "txtStart", "txtEnd", "txtNotes", "txtOfficialComments", "txtOfficialName", "txtRankRate", "txtDesignator", "txtUnitUIC", "txtShipStation", "cboTrait33", "cboTrait34", "cboTrait35", "cboTrait36", "cboTrait37", "cboTrait38", "cboTrait39", "cboObservationStatus", "cboPromotionRecommendation", "cboPromotionStatus", "txtReportingSeniorName", "txtReportingSeniorGrade", "txtReportingSeniorDesignator", "txtReportingSeniorTitle", "txtReportingSeniorUIC", _
        "txtOccasionPeriodic", "txtOccasionDetachment", "txtOccasionFrocking", "txtOccasionSpecial", "txtReportRegular", "txtReportConcurrent", "txtReportOperationalCommander", "txtPrimaryDuty", "txtDutyDescription", _
        "txtMemberSSN", "txtMemberDoDID", "txtComponentActive", "txtComponentTAR", "txtComponentInactive", "txtComponentATADSW", "txtDateReported", _
        "txtCounselingDate", "txtCounselorName", "txtReadinessSummary", "txtReadinessDetail", "txtReadinessDate", "txtBilletSubcategory", _
        "txtAssignmentRecommendationA", "txtAssignmentRecommendationB", "txtRetentionYes", "txtRetentionNo", "txtStatementYes", "txtStatementNo", _
        "txtReportingSeniorSSN", "txtReportingSeniorDoDID", "txtReportingSeniorAddress1", "txtReportingSeniorAddress2", "txtReportingSeniorCity", "txtReportingSeniorState", "txtReportingSeniorZip", "txtReportingSeniorPhone", "txtReportingSeniorDSN", "txtRaterName", "txtRaterDate", "txtSeniorRaterName", "txtSeniorRaterDate", _
        "txtMemberFirstName", "txtMemberMiddleInitial", "txtMemberLastName", "txtMemberSuffix", "txtReportingSeniorLastName", "txtReportingSeniorFirstInitial", "txtReportingSeniorMiddleInitial", "txtCounselorLastName", "txtCounselorFirstInitial", "txtCounselorMiddleInitial", "txtCommandEmployment")
End Function
' Return all 79 pending fields, including active text and the visible surface buffer.
' This reads the draft; it does not create a saved revision.
Public Function EditorPendingFields() As Object
    Dim fields As Object, names As Variant, keys As Variant, i As Long
    Set fields = CreateObject("Scripting.Dictionary")
    names = ControlsList(): keys = DraftFieldKeys()
    For i = 0 To 78
        fields.Add CStr(keys(i)), PendingValue(Forms!frmEditor.Controls(names(i)))
    Next
    Set EditorPendingFields = fields
End Function
' Read the canonical pending control value, including text before focus leaves the field.
' Normalize dates so dirty comparisons and recovery use a stable representation.
Private Function PendingValue(ByVal control As Control) As Variant
    Dim value As Variant
    value = control.Value
    'Capture the text currently being typed, even before focus leaves the textbox.
    On Error Resume Next
    If Screen.ActiveControl.Name = control.Name And Screen.ActiveForm.Name = "frmEditor" Then value = control.Text
    On Error GoTo 0
    value = UXSurfacePending(control.Name, value)
    If VarType(value) = vbDate Then value = Format$(value, "yyyy-mm-dd")
    PendingValue = value
End Function
' Start a fresh recovery session and capture its saved baseline; reset timer/context caches.
Public Function BeginEditorRecovery() As Boolean
    Call UXLiveReset(): Call UXEditorContextChanged()
    checkpointSignature = "": sessionKey = NewIdentity()
    Call ResetEditorBaseline()
    Forms!frmEditor!lblRecovery.Caption = "Recovery: no unsaved edits to checkpoint."
    BeginEditorRecovery = True
End Function
' Capture editor values and the saved revision used by dirty and stale-save guards.
Public Function ResetEditorBaseline() As Boolean
    Dim names As Variant, i As Long
    names = ControlsList()
    For i = 0 To 78: baseline(i) = PendingValue(Forms!frmEditor.Controls(names(i))): Next
    baseRevision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)), 0)
    ResetEditorBaseline = True
End Function
' Reject a save against a newer revision. One user can still change a report through
' another form or import while this editor holds pending values.
Public Function GuardEditorBaseline() As Boolean
    Dim id As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If id > 0 Then
        If Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> baseRevision Then Err.Raise vbObjectError + 900, , "A newer revision exists. Checkpoint these edits, then recover them as a new draft to avoid overwriting it."
    End If
    GuardEditorBaseline = True
End Function
' Compare pending values with the editor baseline, preserving the distinction between
' Null and an empty string; return False when the editor is closed.
Public Function EditorHasPending() As Boolean
    Dim names As Variant, i As Long, value As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Exit Function
    names = ControlsList()
    For i = 0 To 78
        value = PendingValue(Forms!frmEditor.Controls(names(i)))
        If IsNull(value) <> IsNull(baseline(i)) Or StrComp(Nz(value, ""), Nz(baseline(i), ""), vbBinaryCompare) <> 0 Then EditorHasPending = True: Exit Function
    Next
End Function
' Delete only this editor session's checkpoint and invalidate its successful-write signature.
Public Function DiscardEditorCheckpoint() As Boolean
    If Len(sessionKey) > 0 Then CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE SessionGuid='" & sessionKey & "'", dbFailOnError
    checkpointSignature = "": DiscardEditorCheckpoint = True
End Function
' Persist raw pending edits atomically without creating a report revision. Skip an
' unchanged successful checkpoint, recreate missing rows, and retain the prior checkpoint
' on failure; update the signature only after commit.
Public Function CheckpointEditor() As Boolean
    Dim names As Variant, fields As Variant, current(0 To 78) As Variant, i As Long, changed As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, trans As Boolean, signature As String
    On Error GoTo Failed
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Or Len(sessionKey) = 0 Then Exit Function
    names = ControlsList(): fields = Array("FormKind", "DraftTitle", "MemberDisplay", "StartEntry", "EndEntry", "DraftNotes", "OfficialComments", "OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation", "Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39", "ObservationStatus", "PromotionRecommendation", "PromotionStatus", "ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", _
        "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "PrimaryDuty", "DutyDescription", _
        "MemberSSN", "MemberDoDID", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "DateReported", _
        "CounselingDate", "CounselorName", "ReadinessSummary", "ReadinessDetail", "ReadinessDate", "BilletSubcategory", _
        "AssignmentRecommendationA", "AssignmentRecommendationB", "RetentionYes", "RetentionNo", "StatementYes", "StatementNo", _
        "ReportingSeniorSSN", "ReportingSeniorDoDID", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "RaterName", "RaterDate", "SeniorRaterName", "SeniorRaterDate", _
        "MemberFirstName", "MemberMiddleInitial", "MemberLastName", "MemberSuffix", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial", "CounselorLastName", "CounselorFirstInitial", "CounselorMiddleInitial", "CommandEmployment")
    For i = 0 To 78
        current(i) = PendingValue(Forms!frmEditor.Controls(names(i)))
        signature = signature & IIf(IsNull(current(i)), "N", "V") & Len(Nz(current(i), "")) & ":" & Nz(current(i), "")
        If IsNull(current(i)) <> IsNull(baseline(i)) Or StrComp(Nz(current(i), ""), Nz(baseline(i), ""), vbBinaryCompare) <> 0 Then changed = True
    Next
    If Not changed Then
        If DCount("*", "RecoveryDraft", "SessionGuid='" & sessionKey & "'") > 0 Then
            Call DiscardEditorCheckpoint()
            If CurrentProject.AllForms("frmHome").IsLoaded Then Call LoadHome()
        End If
        Forms!frmEditor!lblRecovery.Caption = "Recovery: no unsaved edits to checkpoint."
        CheckpointEditor = True: Exit Function
    End If
    signature = sessionKey & "|" & baseRevision & "|" & Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0) & "|" & signature
    If signature = checkpointSignature Then
        If DCount("*", "RecoveryDraft", "SessionGuid='" & sessionKey & "'") > 0 Then CheckpointEditor = True: Exit Function
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RecoveryDraft WHERE SessionGuid='" & sessionKey & "'", dbOpenDynaset)
    If r.EOF Then r.AddNew: r!SessionGuid = sessionKey Else r.Edit
    r!AuthorProfileID = DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1")
    If Nz(Forms!frmEditor!txtReportID, 0) > 0 Then r!ReportID = Forms!frmEditor!txtReportID Else r!ReportID = Null
    If baseRevision > 0 Then r!BaseRevisionID = baseRevision Else r!BaseRevisionID = Null
    For i = 0 To 78: r.Fields(fields(i)) = current(i): Next
    r!CheckpointAt = Now(): r.Update: r.Close
    w.CommitTrans: trans = False: checkpointSignature = signature
    Forms!frmEditor!lblRecovery.Caption = "Recovery checkpoint: " & Format$(Now(), "h:nn:ss AM/PM") & " - edits are still unsaved. Use Save draft."
    If CurrentProject.AllForms("frmHome").IsLoaded Then Call LoadHome()
    CheckpointEditor = True: Exit Function
Failed:
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    Forms!frmEditor!lblRecovery.Caption = "Recovery failed - latest edits are not checkpointed. Pending values remain here; use Save draft."
End Function
' List checkpoints owned by the active contributor, newest first.
Public Function LoadRecoveries() As Boolean
    Dim pid As Long
    pid = Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)
    Forms!frmRecovery!lstRecoveries.RowSource = "SELECT RecoveryDraftID, DraftTitle, FormKind, CheckpointAt FROM RecoveryDraft WHERE AuthorProfileID=" & pid & " ORDER BY CheckpointAt DESC"
    Forms!frmRecovery!lstRecoveries.Requery
    LoadRecoveries = True
End Function
' Restore a contributor-owned checkpoint into the unbound editor. Reject changed or
' locked saved baselines unless recovering as a new draft; saving remains explicit.
Public Function RestoreRecoveryDraft(ByVal recoveryID As Long, Optional ByVal asNewDraft As Boolean = False) As Boolean
    Dim r As DAO.Recordset, id As Long, names As Variant, fields As Variant, i As Long
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 901, , "Close the current editor before recovering another draft."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RecoveryDraft WHERE RecoveryDraftID=" & recoveryID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 902, , "Checkpoint no longer exists."
    If r!AuthorProfileID <> Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0) Then Err.Raise vbObjectError + 903, , "This checkpoint belongs to another contributor profile."
    id = Nz(r!ReportID, 0)
    If Not asNewDraft And id > 0 Then
        If DCount("*", "Report", "ReportID=" & id) = 0 Or Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> Nz(r!BaseRevisionID, 0) Then Err.Raise vbObjectError + 904, , "The saved report has changed. Use Recover as new draft."
        If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), False) Or Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then Err.Raise vbObjectError + 905, , "The report is archived or locked. Use Recover as new draft."
    End If
    If asNewDraft Then id = 0
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    sessionKey = r!SessionGuid
    names = ControlsList(): fields = Array("FormKind", "DraftTitle", "MemberDisplay", "StartEntry", "EndEntry", "DraftNotes", "OfficialComments", "OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation", "Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39", "ObservationStatus", "PromotionRecommendation", "PromotionStatus", "ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", _
        "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "PrimaryDuty", "DutyDescription", _
        "MemberSSN", "MemberDoDID", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "DateReported", _
        "CounselingDate", "CounselorName", "ReadinessSummary", "ReadinessDetail", "ReadinessDate", "BilletSubcategory", _
        "AssignmentRecommendationA", "AssignmentRecommendationB", "RetentionYes", "RetentionNo", "StatementYes", "StatementNo", _
        "ReportingSeniorSSN", "ReportingSeniorDoDID", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "RaterName", "RaterDate", "SeniorRaterName", "SeniorRaterDate", _
        "MemberFirstName", "MemberMiddleInitial", "MemberLastName", "MemberSuffix", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial", "CounselorLastName", "CounselorFirstInitial", "CounselorMiddleInitial", "CommandEmployment")
    For i = 0 To 78: Forms!frmEditor.Controls(names(i)) = r.Fields(fields(i)).Value: Next
    Forms!frmEditor!lblSaved.Caption = "Recovered pending edits. Use Save draft to commit."
    Forms!frmEditor!lblRecovery.Caption = "Recovered checkpoint " & r!CheckpointAt & " - pending edits, not a saved revision."
    r.Close: Call InitializeTraitEditor(): Call RefreshNarrativeLabel(): Call UXRefreshEditor()
    Forms!frmEditor.SetFocus: RestoreRecoveryDraft = True
End Function
' Restore the selected checkpoint and present recoverable failures to the user.
Public Function RecoverSelectedUI(Optional ByVal asNewDraft As Boolean = False) As Boolean
    On Error GoTo Failed
    RecoverSelectedUI = RestoreRecoveryDraft(Nz(Forms!frmRecovery!lstRecoveries, 0), asNewDraft)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Recover pending edits"
End Function
' Confirm permanent checkpoint removal, restricted to the active contributor and a closed editor.
Public Function DeleteRecoveryUI() As Boolean
    Dim id As Long
    id = Nz(Forms!frmRecovery!lstRecoveries, 0)
    If id = 0 Then Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then MsgBox "Close the editor first.", vbInformation: Exit Function
    If MsgBox("Discard this recovery checkpoint permanently?", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE RecoveryDraftID=" & id & " AND AuthorProfileID=" & Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0), dbFailOnError
    Call LoadRecoveries(): DeleteRecoveryUI = True
End Function
