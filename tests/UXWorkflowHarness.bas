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
    Dim beforePrompts As Long, stage As String, template As Long, route As Long, otherRoute As Long, recovery As Long
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
    f!uxFieldSurface.SetFocus: s!txtOfficialName.SetFocus: s!txtOfficialName.Text = "UNSAVED, NAME"
    Check EditorHasPending(), "Typing in active field not dirty"
    Check DCount("*", "RecoveryDraft") = 0, "Unexpected checkpoint before first timer"
    NativeClose "Save"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Automatic native save did not close"
    Check DLookup("OfficialName", "Report", "ReportID=" & id) = "UNSAVED, NAME", "Automatic close lost active text"
    Check prompts = 0, "Automatic close displayed leave popup"
    stage = "failed save on native close"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Set f = Forms!frmEditor
    f!cmdClose.SetFocus: f!txtMemberSSN = "123": UXSurfaceRefresh
    revision = CurrentRevision(id): NativeClose "Save"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Successful checkpoint trapped invalid draft on native close"
    recovery = DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)
    Check CurrentRevision(id) = revision And DLookup("MemberSSN", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "123", "Native close lost raw value or changed saved revision"
    Check RestoreRecoveryDraft(recovery), "Checkpointed invalid draft could not recover"
    Set f = Forms!frmEditor
    f!txtTitle = "Rejected checkpoint": UXSurfaceRefresh
    CurrentDb.TableDefs("RecoveryDraft").Fields("DraftTitle").ValidationRule = "<> 'Rejected checkpoint'"
    NativeClose "Save"
    Check CurrentProject.AllForms("frmEditor").IsLoaded And f!txtTitle = "Rejected checkpoint", "Failed checkpoint closed or lost the pending buffer"
    CurrentDb.TableDefs("RecoveryDraft").Fields("DraftTitle").ValidationRule = ""
    f!txtTitle = "Close test": UXSurfaceRefresh
    Check EditorHasPending(), "Failed save cleared dirty state"
    Check Not UXLeaveEditor("Cancel") And CurrentProject.AllForms("frmEditor").IsLoaded, "Explicit keep editing did not cancel"
    stage = "save and close"
    f!txtMemberSSN = "000000099": UXSurfaceRefresh
    NativeClose "Save"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Saved native close stayed open"
    Check CurrentRevision(id) <> revision, "Native save did not commit"
    Check DLookup("OfficialName", "Report", "ReportID=" & id) = "UNSAVED, NAME", "Native save lost typed value"
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
    stage = "confirmed shared field edits"
    groupID = CreateSummaryGroup("Workflow group", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = groupID: UXOpenGroup True
    Call UXStartGroupBulk(): Set f = Forms!frmUXBulk
    Check Not UXHasControl(f, "cmdApply") And Not UXHasControl(f, "cmdPreview"), "Preview controls retained"
    revision = CurrentRevision(id): f!uxValueShipStation = "CANCELLED"
    Check Not UXConfirmBlockChange("ShipStation", vbNo), "Cancel applied"
    Check CurrentRevision(id) = revision, "Cancel created revision"
    f!uxValueShipStation = "EDIT B"
    Check UXConfirmBlockChange("ShipStation", vbYes), "Confirmed shared edit failed"
    Check DLookup("ShipStation", "Report", "ReportID=" & id) = "EDIT B", "Confirmed value missing"
    Check CurrentRevision(id) <> revision And Not UXGroupBulkPending(), "Confirmed edit remained staged"
    Check UXBulkSection(0), "Section navigation failed after save"
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
    Set f = UXReviewForm()
    Check CLng(f!txtReportID) = otherID, "Wrong review report"
    Check f!cboRun.ListCount = 1 And CLng(f!cboRun.Column(0, 0)) = otherRoute, "Unrelated route exposed"
    Check InStr(f!lblHeading.Caption, "Second member") > 0, "Review context misleading"
    Check InStr(f!lblSaved.Caption, "saved revision 1") > 0, "Review shows internal revision ID instead of report revision number"
    Check InStr(f!lblSaved.Caption, "2030") = 0 And InStr(f!lblSaved.Caption, "2026") > 0, "Saved-version header showed pending period"
    Check Nz(f!txtGroupID, 0) = 0, "Report review retained unrelated group context"
    Check UXSelectSection(0), "Return from embedded review failed"
    Check UXLeaveEditor("Discard"), "Review editor cleanup failed"
    Forms!frmUXGroups!lstGroups = groupID: Check UXOpenGroup(), "Return to group failed"
    Check CurrentProject.AllForms("frmUXComparison").IsLoaded, "Group members not restored"
    TestUXWorkflows = "PASS: " & checks & " native-close, direct shared-field and report-review assertions": Exit Function
Failed:
    TestUXWorkflows = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
Public Function PrepareForcedQuit(ByVal invalid As Boolean, Optional ByVal kind As String = "EVAL") As Boolean
    Dim id As Long, block As Long, entry As Long
    ConfigureWorkspace "SYNTHETIC QUIT TEST", "Quit tester"
    id = CommitDraft(0, kind, "Forced quit", "Quit member", Null, Null, "", "Test", "")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor.TimerInterval = 0
    Forms!frmEditor!txtOfficialName = "SAVE, QUIT"
    Forms!frmEditor!txtMemberSSN = IIf(invalid, "123", "000000777")
    If kind = "CHIEFEVAL" Then
        For block = 33 To 39
            For entry = 1 To 3
                Forms!frmEditor.Controls("txtTrait" & block & "Comment" & entry) = "QUIT CHIEF " & block & "." & entry
            Next
        Next
    End If
    PrepareForcedQuit = True
End Function
Public Function VerifyChiefRecovery() As String
    Dim block As Long, entry As Long, recovery As Long
    On Error GoTo Failed
    BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
    recovery = DMin("RecoveryDraftID", "RecoveryDraft")
    If Not RestoreRecoveryDraft(recovery) Then Err.Raise 5, , "Chief recovery did not reopen"
    For block = 33 To 39
        For entry = 1 To 3
            If Forms!frmEditor.Controls("txtTrait" & block & "Comment" & entry) <> "QUIT CHIEF " & block & "." & entry Then Err.Raise 5, , "Reopened Chief recovery changed a comment"
        Next
    Next
    If Forms!frmEditor!txtMemberSSN <> "123" Or Not EditorHasPending() Then Err.Raise 5, , "Reopened invalid pending values were coerced"
    If Not UXLeaveEditor("Discard") Then Err.Raise 5, , "Explicit fixture discard failed"
    VerifyChiefRecovery = "PASS: all 21 Chief comments restored in a new Access session with invalid pending data": Exit Function
Failed:
    VerifyChiefRecovery = "FAIL: " & Err.Number & " " & Err.Description
End Function
