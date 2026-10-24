Attribute VB_Name = "LiveMembersHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8010, , message
 checks = checks + 1
End Sub
Public Function TestLiveMembers() As String
 Dim kind As Variant, grade As String, id As Long, other As Long, groupID As Long, f As Form, rows As Form, changed As Object, revision As Long, firstSet As Long, beforeCount As Long, proposal As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC LIVE MEMBERS", "Tester"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  grade = IIf(kind = "EVAL", "E5", IIf(kind = "CHIEFEVAL", "E7", "O3"))
  id = CommitDraft(0, CStr(kind), "Synthetic", "Synthetic", Null, Null, "", "Test", Null, , Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"), Array("Observed", "Promotable"))
  other = DuplicateDraft(id): groupID = CreateSummaryGroup("Live " & kind, CStr(kind), grade)
  AddGroupMember groupID, id, grade: AddGroupMember groupID, other, grade
  beforeCount = DCount("*", "GroupSnapshot")
  Check UXShowGroupMembers(groupID), "Live open failed": Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
  Check f!txtMeeting = 0 And f!uxGroupContext = groupID, "Live group context missing"
  Check DCount("*", "GroupSnapshot") = beforeCount, "Opening Members wrote snapshot"
  rows!Trait33 = "1.0": Check UXGridEdit("Trait33"), "Live save failed: " & f!lblStatus.Caption
  rows!AssignmentRecommendationA = "STAFF": Check UXGridEdit("AssignmentRecommendationA"), "Career save failed"
  Check DLookup("Trait33", "Report", "ReportID=" & id) = "1.0", "Trait not saved"
  Check DCount("*", "GroupSnapshot") = beforeCount And f!txtMeeting = 0, "Live save wrote snapshot"
  revision = CurrentRevision(id): Set changed = CreateObject("Scripting.Dictionary"): changed.Add "Trait34", "3.0"
  CommitSavedFields id, revision, changed, "Other local action"
  Check UXGridLockReason(id, revision) = "", "Row focus ran authoritative revision check"
  rows!Trait35 = "2.0": Check Not UXGridEdit("Trait35"), "Stale save accepted"
  Check DLookup("Trait35", "Report", "ReportID=" & id) = "4.0", "Stale save overwrote report"
  Check UXCloseComparison(), "Close live failed": Check UXShowGroupMembers(groupID), "Reopen live failed"
  Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
  Check rows!Trait34 = "3.0", "Reopen did not read latest saved values"
  CurrentDb.Execute "DELETE FROM GroupMember WHERE GroupID=" & groupID & " AND ReportID=" & id, dbFailOnError
  rows!Trait35 = "2.0": Check Not UXGridEdit("Trait35"), "Removed member save accepted"
  AddGroupMember groupID, id, grade
  Check UXCloseComparison(), "Close failed"
  DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
  Forms!frmMajorWorkbench!cboGroup = groupID: Forms!frmMajorWorkbench!uxMode = "Comparison"
  Check UXOpenComparison(), "Historical comparison failed"
  Set f = Forms!frmUXComparison: firstSet = f!txtMeeting
  Check firstSet > 0 And DCount("*", "ComparisonMember", "SetID=" & firstSet) = 2, "Historical baseline missing"
  ProposeMeetingField firstSet, id, "Trait36", "3.0"
  proposal = ProposeMeetingField(firstSet, id, "Title", "Synthetic non-grid proposal")
  Check UXGridLoad(), "Non-grid proposal broke the trait grid"
  f!lstProposals = proposal: Check UXGridDetail(), "Non-grid proposal details unavailable"
  Check InStr(f!uxProposalDetail, "Synthetic non-grid proposal") > 0, "Non-grid proposal was hidden or truncated"
  Check DLookup("Title", "Report", "ReportID=" & id) = "Synthetic", "Preview committed a non-grid proposal"
  revision = CurrentRevision(id): changed.RemoveAll: changed.Add "Trait37", "3.0": CommitSavedFields id, revision, changed, "Make proposal stale"
  Check RejectProposal(DMax("ProposalID", "MeetingProposal", "SetID=" & firstSet)), "Stale historical proposal accepted"
  Check UXCloseComparison(), "Historical close failed"
  DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
 Next
 ' Explicit staging in Members creates the baseline only when requested.
 Check UXShowGroupMembers(groupID), "Lazy proposal open failed": Set f = Forms!frmUXComparison
 Check UXSelectComparisonMember(), "Select proposal member failed"
 f!cboTrait33 = "2.0": Check UXStageComparison(), "Lazy stage failed: " & f!lblStatus.Caption
 Check f!txtMeeting > 0 And DCount("*", "MeetingProposal", "SetID=" & f!txtMeeting & " AND ProposalState='Proposed'") = 1, "Lazy proposal not staged"
 TestLiveMembers = "PASS: " & checks & " live navigation/save, authoritative guards, three report types, historical conflict and lazy proposal assertions"
 Exit Function
Failed:
 TestLiveMembers = "FAIL: " & Err.Number & " " & Err.Description
End Function
Private Function RejectProposal(ByVal id As Long) As Boolean
 On Error GoTo Rejected
 ApplyMeetingProposals CStr(id)
 Exit Function
Rejected:
 RejectProposal = True
End Function
