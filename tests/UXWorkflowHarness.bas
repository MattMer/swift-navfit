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
Private Function LibraryContains(ByVal reportID As Long) As Boolean
    Dim i As Long
    For i = 1 To Forms!frmLibrary!lstReports.ListCount - 1
        If CLng(Forms!frmLibrary!lstReports.Column(0, i)) = reportID Then LibraryContains = True: Exit Function
    Next
End Function
Private Function RejectSavedFilter(ByVal id As Long, ByVal version As Long, ByVal name As String, ByVal values As Object) As Boolean
    On Error GoTo Rejected
    Call SaveReportFilter(id, version, name, values)
    Exit Function
Rejected:
    RejectSavedFilter = True
End Function
Public Function TestReportFilters() As String
    Dim f As Access.Form, first As Long, outside As Long, undated As Long, saved As Long, owner As Long, other As Long, group As Long
    Dim values As Object, before As Long, prior As String, rejected As Boolean, stage As String, otherSaved As Long
    On Error GoTo Failed
    before = checks: owner = MajorAuthor()
    stage = "search"
    Check ReportSearchMatches("Taylor, Jamie", Null, "Jamie Taylor", Null), "Name order or comma search failed"
    Check ReportSearchMatches("oneil smith", "", "O'Neil-Smith", ""), "Punctuation search failed"
    Check Not ReportSearchMatches("taylor zzz", Null, "Taylor Jamie", Null), "Search accepted a missing token"
    Check ReportSearchMatches("O'BRIEN", Null, "O'Brien", Null), "Apostrophe search failed"
    first = CommitDraft(0, "EVAL", "P01 boundary", "Jamie O'Neil-Smith", #1/1/2026#, #9/30/2026#, "", "Filter test")
    outside = CommitDraft(0, "EVAL", "P01 outside", "Jamie O'Neil-Smith", #1/1/2025#, #12/31/2025#, "", "Filter test")
    undated = CommitDraft(0, "EVAL", "P01 undated", "Jamie O'Neil-Smith", Null, Null, "", "Filter test")
    Set values = CreateObject("Scripting.Dictionary"): values.Add "RankRate", "IT2"
    Call CommitSavedFields(first, CurrentRevision(first), values, "P01 rank filter fixture")
    Check UXNavigate("frmLibrary"), "Reports did not open"
    Set f = Forms!frmLibrary: Check UXClearReportFilters(), "Clear filters failed"
    f!txtSearch = "SMITH, JAMIE ONEIL": Check RefreshLibrary() And LibraryContains(first), "Reordered name did not match in Reports"
    f!txtPeriodFrom = #9/30/2026#: f!txtPeriodThrough = #9/30/2026#
    Check RefreshLibrary() And LibraryContains(first) And Not LibraryContains(outside) And Not LibraryContains(undated), "Inclusive overlap or null-period filter failed"
    prior = f!lstReports.RowSource: f!txtPeriodFrom = #1/1/2027#
    Check Not RefreshLibrary() And f!lstReports.RowSource = prior, "Reversed period changed the report list"
    f!txtPeriodFrom = #9/30/2026#
    stage = "saved sets"
    group = CreateSummaryGroup("P01 saved scope", "EVAL", "E5"): AddGroupMember group, first, "E5"
    Check RefreshFolders(), "Group choices did not refresh"
    UXMemberRankChoices f, "Report"
    f!uxTypeFilter = "EVAL": f!uxStatusFilter = "Draft": f!cboRankFilter = "E5"
    f!lstFolders = group: f!uxFolderFilter = group
    Check UXSaveReportFilter("P01 saved O'Brien", True), "Named filter save failed: " & f!lblCount.Caption
    saved = CLng(f!cboSavedFilter): Set values = ReportFilterCriteria()
    Check UXClearReportFilters() And IsNull(f!txtPeriodFrom) And IsNull(f!cboSavedFilter), "Clear filters retained a period or named selection"
    f!cboSavedFilter = saved
    Check UXLoadReportFilter() And LibraryContains(first) And Not LibraryContains(outside), "Named filter did not restore the list: " & f!lblCount.Caption
    Check f!txtPeriodFrom = #9/30/2026# And CLng(f!lstFolders) = group, "Named period/group scope lost"
    Check f!uxTypeFilter = "EVAL" And f!uxStatusFilter = "Draft" And f!cboRankFilter = "E5" And f!cboArchive = "Active", "Named type/status/rank/archive criteria lost"
    Check RejectSavedFilter(0, 0, "P01 saved O'Brien", values), "Duplicate name accepted"
    Check SaveReportFilter(saved, 1, "P01 saved O'Brien", values) = saved, "Named filter update failed"
    Check RejectSavedFilter(saved, 1, "P01 saved O'Brien", values), "Stale named-filter update accepted"
    stage = "contributor isolation"
    other = AddContributor("Synthetic filter owner")
    Check SwitchContributor(other), "Contributor switch failed"
    Check f!cboSavedFilter.ListCount = 0 And IsNull(f!cboSavedFilter), "Contributor switch exposed stale saved sets"
    Check RejectSavedFilter(saved, 2, "P01 saved O'Brien", values), "Another contributor replaced the saved set"
    otherSaved = SaveReportFilter(0, 0, "P01 saved O'Brien", values)
    Check otherSaved <> saved, "Contributor could not use the same filter name"
    Check DeleteReportFilter(otherSaved, 1), "Own named-filter delete failed"
    Check SwitchContributor(owner), "Owner switch failed"
    Check f!cboSavedFilter.ListCount = 1 And IsNull(f!cboSavedFilter), "Owner filter choices did not refresh on switch"
    f!cboSavedFilter = saved: Check UXLoadReportFilter(), "Owner could not reload saved set"
    On Error Resume Next
    Call DeleteReportFilter(saved, 1): rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check rejected And DCount("*", "SavedReportFilter", "FilterID=" & saved) = 1, "Stale filter delete removed the set"
    Check SaveReportFilter(saved, 2, "P01 saved O'Brien", values) = saved, "Concurrent filter update failed"
    Check Not UXDeleteReportFilter(True) And DCount("*", "SavedReportFilter", "FilterID=" & saved) = 1, "Stale UI deleted a newer saved set"
    Check UXLoadReportFilter(), "Saved set did not refresh after a conflict"
    Check UXSaveReportFilter("P01 saved O'Brien", True), "Confirmed UI replacement failed"
    prior = f!lstReports.RowSource
    CurrentDb.Execute "UPDATE SummaryGroup SET IsArchived=True WHERE GroupID=" & group, dbFailOnError
    Check Not UXLoadReportFilter() And f!lstReports.RowSource = prior, "Archived group silently broadened a saved filter"
    CurrentDb.Execute "UPDATE SummaryGroup SET IsArchived=False WHERE GroupID=" & group, dbFailOnError
    Check UXLoadReportFilter(), "Restored group scope could not reload"
    Check UXClearReportFilters(), "Clear before saving an empty search failed"
    Check UXSaveReportFilter("P01 temporary", True), "Empty-search filter save failed"
    Check UXDeleteReportFilter(True), "Named filter UI delete failed: " & f!lblCount.Caption
    f!cboSavedFilter = saved: Check UXLoadReportFilter(), "Final saved set did not reload"
    TestReportFilters = "PASS: " & (checks - before) & " report search, period, named filter, owner and stale-write checks": Exit Function
Failed:
    TestReportFilters = "FAIL report filters at " & stage & ": " & Err.Number & " " & Err.Description
End Function
Public Function VerifyReportFilterRecovery() As String
    Dim id As Long
    On Error GoTo Failed
    Call SwitchContributor(CLng(DMin("UserProfileID", "UserProfile")))
    id = DLookup("FilterID", "SavedReportFilter", "FilterName='P01 saved O''Brien'")
    DoCmd.OpenForm "frmLibrary"
    Forms!frmLibrary!cboSavedFilter = id
    If Not UXLoadReportFilter() Then Err.Raise 5, , Forms!frmLibrary!lblCount.Caption
    If Forms!frmLibrary!txtPeriodFrom <> #9/30/2026# Or Forms!frmLibrary!txtSearch <> "SMITH, JAMIE ONEIL" Or Nz(Forms!frmLibrary!lstFolders, 0) <= 0 Then Err.Raise 5, , "Saved criteria lost on reopen"
    VerifyReportFilterRecovery = "PASS: saved report filter restores after a fresh Access session": Exit Function
Failed:
    VerifyReportFilterRecovery = "FAIL: " & Err.Number & " " & Err.Description
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
