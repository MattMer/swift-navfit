Attribute VB_Name = "UXWorkflowHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private nextDecision As String
Private prompts As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 5300, , message
    checks = checks + 1
End Sub
Public Function WorkflowDecisionTimer() As Boolean
    Forms!frmUXDecision.TimerInterval = 0
    prompts = prompts + 1
    WorkflowDecisionTimer = UXDecide(nextDecision)
End Function
Private Sub NativeClose(ByVal decision As String)
    nextDecision = decision
    On Error Resume Next
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Err.Clear
    On Error GoTo 0
End Sub
Public Function TestUXWorkflows() As String
    Dim id As Long, otherID As Long, thirdID As Long, groupID As Long, revision As Long, f As Access.Form, s As Access.Form
    Dim beforePrompts As Long, stage As String, template As Long, route As Long, otherRoute As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC WORKFLOW TEST", "Workflow tester"
    ' Script only the dialog choice in this scratch fixture; exercise the real Unload event.
    DoCmd.OpenForm "frmUXDecision", acDesign
    With Forms!frmUXDecision
        .TimerInterval = 1: .OnTimer = "=WorkflowDecisionTimer()"
    End With
    DoCmd.Close acForm, "frmUXDecision", acSaveYes
    id = CommitDraft(0, "EVAL", "Close test", "First member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    stage = "native automatic save"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
    Check f.OnUnload = "[Event Procedure]", "Unload guard missing"
    f!uxFieldSurface.SetFocus: s!txtOfficialName.SetFocus: s!txtOfficialName.Text = "UNSAVED NAME"
    Check EditorHasPending(), "Typing in active field not dirty"
    Check DCount("*", "RecoveryDraft") = 0, "Unexpected checkpoint before first timer"
    NativeClose "Save"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Automatic native save did not close"
    Check DLookup("OfficialName", "Report", "ReportID=" & id) = "UNSAVED NAME", "Automatic close lost active text"
    Check prompts = 0, "Automatic close displayed leave popup"
    stage = "failed save on native close"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Set f = Forms!frmEditor
    f!cmdClose.SetFocus: f!txtMemberSSN = "123": UXSurfaceRefresh
    revision = CurrentRevision(id): NativeClose "Save"
    Check CurrentProject.AllForms("frmEditor").IsLoaded, "Failed save closed report"
    Check CurrentRevision(id) = revision And f!txtMemberSSN = "123", "Failed save lost value or changed revision"
    Check EditorHasPending(), "Failed save cleared dirty state"
    Check Not UXLeaveEditor("Cancel") And CurrentProject.AllForms("frmEditor").IsLoaded, "Explicit keep editing did not cancel"
    stage = "save and close"
    f!txtMemberSSN = "000000099": UXSurfaceRefresh
    NativeClose "Save"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Saved native close stayed open"
    Check CurrentRevision(id) <> revision, "Native save did not commit"
    Check DLookup("OfficialName", "Report", "ReportID=" & id) = "UNSAVED NAME", "Native save lost typed value"
    Check DCount("*", "RecoveryDraft") = 0, "Save left pending checkpoint"
    stage = "unchanged close"
    beforePrompts = prompts: DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    NativeClose "Cancel"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded And prompts = beforePrompts, "Unchanged close prompted"
    stage = "discard checkpoint"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtNotes = "Discard me": Check CheckpointEditor(), "Checkpoint failed"
    revision = CurrentRevision(id): Check UXLeaveEditor("Discard"), "Explicit discard failed"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Discard did not close"
    Check CurrentRevision(id) = revision And DCount("*", "RecoveryDraft") = 0, "Discard committed or retained checkpoint"
    stage = "custom leave"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtNotes = "Custom leave": beforePrompts = prompts
    revision = CurrentRevision(id)
    Check UXLeaveEditor(), "Custom automatic save failed"
    Check CurrentRevision(id) <> revision, "Custom leave did not save"
    Check prompts = beforePrompts, "Custom leave asked a second time"
    stage = "bulk invalidation"
    groupID = CreateSummaryGroup("Workflow group", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = groupID: UXOpenGroup
    Call UXStartGroupBulk(): Set f = Forms!frmUXBulk
    Check Not f!cmdApply.Enabled, "Apply enabled without preview"
    f!uxValueShipStation = "PREVIEW A": f!txtReason = "Review changes": UXBulkChanged "ShipStation"
    Check UXPreviewBulk() And f!cmdApply.Enabled, "Preview did not enable Apply"
    f!uxValueShipStation = "EDIT B": UXBulkChanged "ShipStation"
    Check Not f!cmdApply.Enabled And Len(f!lstPreview.RowSource) = 0, "Changed input retained preview"
    revision = CurrentRevision(id)
    Check Not UXApplyBulk() And CurrentRevision(id) = revision, "Stale input applied"
    Check UXPreviewBulk(), "Fresh preview failed"
    f!txtReason = "Different reason"
    Check Not UXApplyBulk() And Not f!cmdApply.Enabled, "Reason change applied old preview"
    Check UXPreviewBulk(), "Reason refresh failed"
    Check UXApplyBulk(), "Reviewed EDIT B apply failed"
    Check DLookup("ShipStation", "Report", "ReportID=" & id) = "EDIT B", "Applied old value"
    Check Not f!cmdApply.Enabled And Not UXApplyBulk(), "Repeated apply allowed"
    f!uxValueShipStation = "Membership change": UXBulkChanged "ShipStation": Check UXPreviewBulk(), "Second preview failed"
    thirdID = CommitDraft(0, "EVAL", "New group member", "Third member", Null, Null, "", "Test", "")
    AddGroupMember groupID, thirdID, "E5": revision = CurrentRevision(id)
    Check Not UXApplyBulk() And CurrentRevision(id) = revision, "Changed membership applied"
    Check Not UXPreviewBulk(), "Old target roster previewed after membership changed"
    Call UXStartGroupBulk(): Check Not f!cmdApply.Enabled, "Reopened bulk retained old batch"
    Check InStr(f!txtReason.AfterUpdate, "UXInvalidateBulkPreview") > 0, "Reason event missing"
    Check InStr(f!txtReason.OnChange, "UXInvalidateBulkPreview") > 0 And InStr(f!uxValueShipStation.OnChange, "UXInvalidateBulkPreview") > 0, "Typing does not invalidate preview"
    DoCmd.Close acForm, "frmUXBulk", acSaveNo
    stage = "report review context"
    otherID = CommitDraft(0, "EVAL", "Outside group", "Second member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    template = CreateRouteTemplate("Workflow route", CStr(MajorAuthor()), "Check saved report")
    route = StartRoute(id, template, groupID): otherRoute = StartRoute(otherID, template, 0)
    Forms!frmMajorWorkbench!cboRun = route
    DoCmd.OpenForm "frmEditor", , , , , , CStr(otherID)
    Forms!frmEditor!txtNotes = "Keep these pending review notes"
    Forms!frmEditor!txtStart = #1/1/2030#
    revision = CurrentRevision(otherID): nextDecision = "Discard"
    Check UXEditorReview("Review"), "Report review did not open"
    Check EditorHasPending() And CurrentRevision(otherID) = revision, "Use saved version discarded or saved pending edits"
    Set f = Forms!frmUXReview
    Check CLng(f!txtReportID) = otherID, "Wrong review report"
    Check f!cboRun.ListCount = 1 And CLng(f!cboRun.Column(0, 0)) = otherRoute, "Unrelated route exposed"
    Check InStr(f!lblHeading.Caption, "Second member") > 0, "Review context misleading"
    Check InStr(f!lblSaved.Caption, "saved revision 1") > 0, "Review shows internal revision ID instead of report revision number"
    Check InStr(f!lblSaved.Caption, "2030") = 0 And InStr(f!lblSaved.Caption, "2026") > 0, "Saved-version header showed pending period"
    Check Nz(f!txtGroupID, 0) = 0, "Report review retained unrelated group context"
    DoCmd.Close acForm, "frmUXReview", acSaveNo
    Check UXLeaveEditor("Discard"), "Review editor cleanup failed"
    Forms!frmUXGroups!lstGroups = groupID: Check UXOpenGroup(), "Return to group failed"
    Check Forms!frmMajorWorkbench!uxMode0.Visible And Forms!frmMajorWorkbench!uxMode3.Visible, "Group navigation not restored"
    TestUXWorkflows = "PASS: " & checks & " native-close, bulk-preview and report-review assertions": Exit Function
Failed:
    TestUXWorkflows = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
Public Function PrepareForcedQuit(ByVal invalid As Boolean) As Boolean
    Dim id As Long
    ConfigureWorkspace "SYNTHETIC QUIT TEST", "Quit tester"
    id = CommitDraft(0, "EVAL", "Forced quit", "Quit member", Null, Null, "", "Test", "")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor.TimerInterval = 0
    Forms!frmEditor!txtOfficialName = "QUIT SAVE NAME"
    Forms!frmEditor!txtMemberSSN = IIf(invalid, "123", "000000777")
    PrepareForcedQuit = True
End Function
