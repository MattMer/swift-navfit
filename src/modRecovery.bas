Attribute VB_Name = "modRecovery"
Option Compare Database
Option Explicit
Private sessionKey As String
Private baseRevision As Long
Private baseline(0 To 67) As Variant
Private Function ControlsList() As Variant
    ControlsList = Array("cboKind", "txtTitle", "txtMember", "txtStart", "txtEnd", "txtNotes", "txtOfficialComments", "txtOfficialName", "txtRankRate", "txtDesignator", "txtUnitUIC", "txtShipStation", "cboTrait33", "cboTrait34", "cboTrait35", "cboTrait36", "cboTrait37", "cboTrait38", "cboTrait39", "cboObservationStatus", "cboPromotionRecommendation", "cboPromotionStatus", "txtReportingSeniorName", "txtReportingSeniorGrade", "txtReportingSeniorDesignator", "txtReportingSeniorTitle", "txtReportingSeniorUIC", _
        "txtOccasionPeriodic", "txtOccasionDetachment", "txtOccasionFrocking", "txtOccasionSpecial", "txtReportRegular", "txtReportConcurrent", "txtReportOperationalCommander", "txtPrimaryDuty", "txtDutyDescription", _
        "txtMemberSSN", "txtMemberDoDID", "txtComponentActive", "txtComponentTAR", "txtComponentInactive", "txtComponentATADSW", "txtDateReported", _
        "txtCounselingDate", "txtCounselorName", "txtReadinessSummary", "txtReadinessDetail", "txtReadinessDate", "txtBilletSubcategory", _
        "txtAssignmentRecommendationA", "txtAssignmentRecommendationB", "txtRetentionYes", "txtRetentionNo", "txtStatementYes", "txtStatementNo", _
        "txtReportingSeniorSSN", "txtReportingSeniorDoDID", "txtReportingSeniorAddress1", "txtReportingSeniorAddress2", "txtReportingSeniorCity", "txtReportingSeniorState", "txtReportingSeniorZip", "txtReportingSeniorPhone", "txtReportingSeniorDSN", "txtRaterName", "txtRaterDate", "txtSeniorRaterName", "txtSeniorRaterDate")
End Function
Private Function PendingValue(ByVal control As Control) As Variant
    Dim value As Variant
    value = control.Value
    'Capture the text currently being typed, even before focus leaves the textbox.
    On Error Resume Next
    If Screen.ActiveControl.Name = control.Name And Screen.ActiveForm.Name = "frmEditor" Then value = control.Text
    On Error GoTo 0
    If VarType(value) = vbDate Then value = Format$(value, "yyyy-mm-dd")
    PendingValue = value
End Function
Public Function BeginEditorRecovery() As Boolean
    sessionKey = NewIdentity()
    Call ResetEditorBaseline()
    BeginEditorRecovery = True
End Function
Public Function ResetEditorBaseline() As Boolean
    Dim names As Variant, i As Long
    names = ControlsList()
    For i = 0 To 67: baseline(i) = PendingValue(Forms!frmEditor.Controls(names(i))): Next
    baseRevision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)), 0)
    ResetEditorBaseline = True
End Function
Public Function GuardEditorBaseline() As Boolean
    Dim id As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If id > 0 Then
        If Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> baseRevision Then Err.Raise vbObjectError + 900, , "A newer revision exists. Checkpoint these edits, then recover them as a new draft to avoid overwriting it."
    End If
    GuardEditorBaseline = True
End Function
Public Function DiscardEditorCheckpoint() As Boolean
    If Len(sessionKey) > 0 Then CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE SessionGuid='" & sessionKey & "'", dbFailOnError
    DiscardEditorCheckpoint = True
End Function
Public Function CheckpointEditor() As Boolean
    Dim names As Variant, fields As Variant, current(0 To 67) As Variant, i As Long, changed As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, trans As Boolean
    On Error GoTo Failed
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Or Len(sessionKey) = 0 Then Exit Function
    names = ControlsList(): fields = Array("FormKind", "DraftTitle", "MemberDisplay", "StartEntry", "EndEntry", "DraftNotes", "OfficialComments", "OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation", "Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39", "ObservationStatus", "PromotionRecommendation", "PromotionStatus", "ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", _
        "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "PrimaryDuty", "DutyDescription", _
        "MemberSSN", "MemberDoDID", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "DateReported", _
        "CounselingDate", "CounselorName", "ReadinessSummary", "ReadinessDetail", "ReadinessDate", "BilletSubcategory", _
        "AssignmentRecommendationA", "AssignmentRecommendationB", "RetentionYes", "RetentionNo", "StatementYes", "StatementNo", _
        "ReportingSeniorSSN", "ReportingSeniorDoDID", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "RaterName", "RaterDate", "SeniorRaterName", "SeniorRaterDate")
    For i = 0 To 67
        current(i) = PendingValue(Forms!frmEditor.Controls(names(i)))
        If IsNull(current(i)) <> IsNull(baseline(i)) Or StrComp(Nz(current(i), ""), Nz(baseline(i), ""), vbBinaryCompare) <> 0 Then changed = True
    Next
    If Not changed Then
        Call DiscardEditorCheckpoint()
        Forms!frmEditor!lblRecovery.Caption = "No pending changes to checkpoint."
        If CurrentProject.AllForms("frmHome").IsLoaded Then Call LoadHome()
        CheckpointEditor = True: Exit Function
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RecoveryDraft WHERE SessionGuid='" & sessionKey & "'", dbOpenDynaset)
    If r.EOF Then r.AddNew: r!SessionGuid = sessionKey Else r.Edit
    r!AuthorProfileID = DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1")
    If Nz(Forms!frmEditor!txtReportID, 0) > 0 Then r!ReportID = Forms!frmEditor!txtReportID Else r!ReportID = Null
    If baseRevision > 0 Then r!BaseRevisionID = baseRevision Else r!BaseRevisionID = Null
    For i = 0 To 67: r.Fields(fields(i)) = current(i): Next
    r!CheckpointAt = Now(): r.Update: r.Close
    w.CommitTrans: trans = False
    Forms!frmEditor!lblRecovery.Caption = "Recovery checkpoint: " & Now() & " (not a saved revision)"
    If CurrentProject.AllForms("frmHome").IsLoaded Then Call LoadHome()
    CheckpointEditor = True: Exit Function
Failed:
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    Forms!frmEditor!lblRecovery.Caption = "Checkpoint failed. Use Save draft to retain these edits."
End Function
Public Function LoadRecoveries() As Boolean
    Dim pid As Long
    pid = Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)
    Forms!frmRecovery!lstRecoveries.RowSource = "SELECT RecoveryDraftID, DraftTitle, FormKind, CheckpointAt FROM RecoveryDraft WHERE AuthorProfileID=" & pid & " ORDER BY CheckpointAt DESC"
    Forms!frmRecovery!lstRecoveries.Requery
    LoadRecoveries = True
End Function
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
        "ReportingSeniorSSN", "ReportingSeniorDoDID", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "RaterName", "RaterDate", "SeniorRaterName", "SeniorRaterDate")
    For i = 0 To 67: Forms!frmEditor.Controls(names(i)) = r.Fields(fields(i)).Value: Next
    Forms!frmEditor!lblSaved.Caption = "Recovered pending edits. Use Save draft to commit."
    Forms!frmEditor!lblRecovery.Caption = "Recovered checkpoint " & r!CheckpointAt
    r.Close: Call InitializeTraitEditor(): Call RefreshNarrativeLabel()
    Forms!frmEditor.SetFocus: RestoreRecoveryDraft = True
End Function
Public Function RecoverSelectedUI(Optional ByVal asNewDraft As Boolean = False) As Boolean
    On Error GoTo Failed
    RecoverSelectedUI = RestoreRecoveryDraft(Nz(Forms!frmRecovery!lstRecoveries, 0), asNewDraft)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Recover pending edits"
End Function
Public Function DeleteRecoveryUI() As Boolean
    Dim id As Long
    id = Nz(Forms!frmRecovery!lstRecoveries, 0)
    If id = 0 Then Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then MsgBox "Close the editor first.", vbInformation: Exit Function
    If MsgBox("Discard this recovery checkpoint permanently?", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE RecoveryDraftID=" & id & " AND AuthorProfileID=" & Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0), dbFailOnError
    Call LoadRecoveries(): DeleteRecoveryUI = True
End Function
