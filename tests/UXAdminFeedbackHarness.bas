Attribute VB_Name = "UXAdminFeedbackHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6950, , message
    checks = checks + 1
End Sub
Public Function TestUXAdminFeedback() As String
    Dim id As Long, groupID As Long, revision As Long, duplicate As Long, kind As Variant, rate As Variant, key As Variant
    Dim f As Access.Form, s As Access.Form, changes As Object, proposal As Long, templateID As Long, runID As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC ADMIN FEEDBACK", "Tester"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = CommitDraft(0, CStr(kind), "Admin test", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
        Check UXSectionName(0) = "Admin data", "Admin section caption incorrect"
        Check s!txtOfficialName.Visible And s!txtReportingSeniorName.Visible, "Admin data did not combine member and senior fields"
        Check Not s!txtMemberFirstName.Visible And Not s!txtReportingSeniorLastName.Visible, "Redundant fields exposed"
        'NAVPERS 1616/27 block 20 is the date and a result list; the other forms keep the summary code.
        If kind = "CHIEFEVAL" Then
            Check s!txtReadinessDetail.Visible And Not s!txtReadinessSummary.Visible, "Chief block 20 is not the date and result"
        Else
            Check Not s!txtReadinessDetail.Visible And s!txtReadinessSummary.Visible, "Readiness result exposed outside the Chief form"
        End If
        Check s!txtCommandEmployment.Visible, "Block 28 absent"
        s!txtCommandEmployment = "Synthetic deployment and command achievement"
        Check UXSurfaceEdited("CommandEmployment"), "Block 28 edit failed"
        revision = CurrentRevision(id)
        Check CheckpointEditor(), "Block 28 checkpoint failed"
        Check DLookup("CommandEmployment", "RecoveryDraft", "ReportID=" & id) = "Synthetic deployment and command achievement", "Checkpoint omitted Block 28"
        Check UXSelectSection(4), "Trait section failed"
        Check f!txtCommandEmployment = "Synthetic deployment and command achievement" And CurrentRevision(id) = revision, "Section navigation lost or saved pending Block 28"
        ' Fixed page layout: shared guidance starts level with block 33 and covers through block 39;
        ' Chief evaluations show per-trait guidance instead.
        If kind = "CHIEFEVAL" Then
            Check s!uxChiefGuidance33.Visible And s!uxChiefGuidance39.Visible, "Chief trait guidance missing"
        Else
            Check s!uxTraitGuidance.Visible And s!uxTraitGuidance.Top = s!cboTrait33.Top And s!uxTraitGuidance.Top + s!uxTraitGuidance.Height >= s!cboTrait39.Top + s!cboTrait39.Height, "Trait guidance does not span the trait list"
        End If
        Check SaveEditor(), "Block 28 save failed"
        Check MajorRevisionValue(CurrentRevision(id), "CommandEmployment") = "Synthetic deployment and command achievement", "Block 28 revision missing"
        duplicate = DuplicateDraft(id)
        Check DLookup("CommandEmployment", "Report", "ReportID=" & duplicate) = "Synthetic deployment and command achievement", "Duplicate lost Block 28"
        Check UXLeaveEditor(), "Editor leave failed"
    Next
    groupID = CreateSummaryGroup("Feedback group", "FITREP", "O3"): AddGroupMember groupID, id, "O3"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
    Check Forms!frmMajorWorkbench!uxMode1.OnClick = "=UXOpenGroupDefaults()", "Defaults requires intermediate navigation"
    Check Forms!frmMajorWorkbench!uxMode3.OnClick = "=UXOpenComparison()", "Comparison requires intermediate navigation"
    Check Forms!frmMajorWorkbench!lstReports.OnDblClick = "=UXGroupRowOpen()", "Member double click missing"
    Forms!frmMajorWorkbench!uxMode = "Comparison"
    Check UXOpenComparison(), "Comparison failed"
    Set f = Forms!frmUXComparison: Set s = f!uxTraitGrid.Form
    Check s!Member.OnClick = "=UXGridOpenReport()", "Comparison name click missing"
    s!AssignmentRecommendationA = "INSTRUCTOR"
    Check UXGridEdit("AssignmentRecommendationA"), "Immediate cell save failed: " & f!lblStatus.Caption
    Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id) = "INSTRUCTOR", "Cell did not save report"
    s!AssignmentRecommendationB = "RECRUITER"
    Check UXGridEdit("AssignmentRecommendationB"), "Repeated cell save failed: " & f!lblStatus.Caption
    proposal = ProposeMeetingField(f!txtMeeting, id, "AssignmentRecommendationA", "STAFF")
    UXRefreshComparisonPreview
    f!lstProposals.Selected(1) = True
    Check UXApplyComparison(), "Reviewed proposal apply failed: " & f!lblStatus.Caption
    Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id) = "STAFF", "Reviewed proposal not committed"
    UXCloseComparison
    templateID = CreateRouteTemplate("Approval test", CStr(MajorAuthor()), "Synthetic")
    runID = StartRoute(id, templateID, groupID)
    DecideRouteStep runID, "Accepted", "Synthetic acceptance"
    ApproveSavedRoute runID, "Synthetic approval"
    Check UXOpenComparison(), "Approved comparison did not open"
    Set f = Forms!frmUXComparison: Set s = f!uxTraitGrid.Form
    revision = CurrentRevision(id)
    s!AssignmentRecommendationA = "OPERATIONS"
    Check Not UXGridEdit("AssignmentRecommendationA", vbNo), "Declining approval reopening saved a change"
    Check CurrentRevision(id) = revision And DLookup("WorkflowState", "Report", "ReportID=" & id) = "Approved", "Declining reopening altered approval"
    s!AssignmentRecommendationA = "OPERATIONS"
    Check UXGridEdit("AssignmentRecommendationA", vbYes), "Approved cell could not reopen and save: " & f!lblStatus.Caption
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft" And DLookup("RunState", "RouteRun", "RunID=" & runID) = "Reopened", "Approval history not retained on reopening"
    Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id) = "OPERATIONS", "Reopened change not persisted"
    UXCloseComparison
    DoCmd.OpenForm "frmMyWork": UXLoadWorkQueue
    Check Forms!frmMyWork!uxContributor = MajorAuthor(), "My Work contributor missing"
    DoCmd.OpenForm "frmSnippets": LoadSnippets
    For Each rate In NavyWriterRates()
        Forms!frmSnippets!cboRateFilter = CStr(rate): LoadSnippets
        Check Forms!frmSnippets!lstSnippets.ListCount > 0, "Empty snippet rate: " & CStr(rate)
    Next
    revision = DCount("*", "NarrativeSnippet"): EnsureStarterSnippets
    Check DCount("*", "NarrativeSnippet") = revision, "Starter population duplicated examples"
    DoCmd.OpenForm "frmUXSeniorEditor", , , , , , "0"
    For Each key In Array("ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial")
        Check Not Forms!frmUXSeniorEditor.Controls("rs" & key).Visible, "Directory exposes redundant name part"
    Next
    Check Not RSSeniorDirty(), "Blank directory has phantom pending values"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    CheckLegacyBlock28
    TestUXAdminFeedback = "PASS: " & checks & " admin, Block 28, comparison and snippet assertions": Exit Function
Failed:
    TestUXAdminFeedback = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Private Sub CheckLegacyBlock28()
    Dim id As Long, revision As Long, receipt As Long, db As DAO.Database, r As DAO.Recordset, values As Object
    id = CommitDraft(0, "EVAL", "Legacy source", "Synthetic legacy member", Null, Null, "")
    revision = CurrentRevision(id): Set db = CurrentDb
    Check DCount("*", "RevisionChange", "RevisionID=" & revision & " AND FieldKey='CommandEmployment'") = 0, "New report logged untouched Block 28"
    Set r = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    r.AddNew: r!ReceiptGuid = NewIdentity(): r!SourceHash = String$(64, "0"): r!ExternalReportID = 42: r!ReportID = id
    r!AuthorProfileID = MajorAuthor(): r!ImportedAt = Now(): r!AdapterVersion = "navfit-full-v1/local-CStr": r!IsUndone = False
    r.Update: r.Bookmark = r.LastModified: receipt = r!ImportReceiptID: r.Close
    Set r = db.OpenRecordset("ImportRawValue", dbOpenDynaset)
    r.AddNew: r!ImportReceiptID = receipt: r!SourceTable = "Folders": r!SourceRow = 42
    r!FieldOrdinal = 1: r!FieldKey = "Achievements": r!DAOType = 12: r!RawValue = "Wrong folder default": r!IsNullValue = False: r.Update
    r.AddNew: r!ImportReceiptID = receipt: r!SourceTable = "Reports": r!SourceRow = 42
    r!FieldOrdinal = 1: r!FieldKey = "Achievements": r!DAOType = 12: r!RawValue = "Original native command achievements": r!IsNullValue = False: r.Update: r.Close
    'A current first revision records Block 28 in its snapshot, so the native value is not projected.
    Check IsNull(StoredBlock28(id)), "New report projected retained native Block 28"
    Set values = SavedFields(id)
    Check IsNull(values("CommandEmployment")), "New report export used retained native Block 28"
    Check IsNull(Block28FromRevision(revision)), "New revision invented projected Block 28 snapshot text"
    db.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & revision & " AND FieldKey='CommandEmployment'", dbFailOnError
    db.Execute "DELETE FROM RevisionChange WHERE RevisionID=" & revision & " AND FieldKey='CommandEmployment'", dbFailOnError
    db.Execute "UPDATE ReportRevision SET SnapshotSchema=23, FieldDefinitionVersion=Null WHERE RevisionID=" & revision, dbFailOnError
    Check StoredBlock28(id) = "Original native command achievements", "Legacy Block 28 projection used wrong source row"
    Set values = SavedFields(id)
    Check values("CommandEmployment") = "Original native command achievements", "Export values omitted legacy Block 28 provenance"
    Check IsNull(DLookup("CommandEmployment", "Report", "ReportID=" & id)), "Read-only projection changed a saved report"
    Check IsNull(Block28FromRevision(revision)), "Old revision invented a recorded Block 28 value"
    Set values = CreateObject("Scripting.Dictionary"): values.Add "Title", "Legacy source renamed"
    Call CommitSavedFields(id, revision, values, "Unrelated saved-field edit")
    Check CurrentRevision(id) <> revision And Nz(StoredBlock28(id), "") = "Original native command achievements", "Unrelated save cleared legacy Block 28"
    revision = CurrentRevision(id)
    Set values = CreateObject("Scripting.Dictionary"): values.Add "CommandEmployment", Null
    Call CommitSavedFields(id, revision, values, "Explicitly clear inherited Block 28")
    Check CurrentRevision(id) <> revision And IsNull(StoredBlock28(id)), "Explicit clear revived prior native Block 28"
    Check DCount("*", "RevisionChange", "RevisionID=" & CurrentRevision(id) & " AND FieldKey='CommandEmployment' AND Not OldIsNull AND NewIsNull") = 1, "Inherited Block 28 clear lost its actual prior value"
    Check DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt & " AND RawValue='Original native command achievements'") = 1, "Explicit clear removed provenance"
    Set values = CreateObject("Scripting.Dictionary"): values.Add "Title", "Legacy clear retained"
    Call CommitSavedFields(id, CurrentRevision(id), values, "Routine save after inherited clear")
    Check IsNull(StoredBlock28(id)), "Routine save revived cleared native Block 28"
End Sub


