Attribute VB_Name = "GroupCorrectnessHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6500, , message
    checks = checks + 1
End Sub
Private Sub SaveFixture(ByVal id As Long, ByVal fields As Object, ByVal reason As String)
    Dim changes As Object, key As Variant
    Set changes = CreateObject("Scripting.Dictionary")
    For Each key In fields.Keys
        If key <> "FormKind" Then changes.Add CStr(key), fields(key)
    Next
    Call CommitSavedFields(id, CurrentRevision(id), changes, reason)
End Sub
Private Function CapacityText(ByVal groupID As Long) As String
    Dim item As Object
    For Each item In ValidateLiveGroup(groupID)
        If item("Code") = "CAPACITY" Then CapacityText = item("Message")
    Next
End Function
Public Function TestGroupCorrectness() As String
    Dim g As Long, id As Long, ids(1 To 3) As Long, i As Long, f As Object, text As String, rev As Long, operations As Long, stage As String, stamp As String, snapshots As Long, overrides As Object, preview As Long, before As String
    On Error GoTo Failed
    stage = "Configure": ConfigureWorkspace "SYNTHETIC GROUP CORRECTNESS", "Group tester"
    stage = "Create group": g = CreateSummaryGroup("Observed E5", "EVAL", "E5")
    For i = 1 To 3
        stage = "Create member": id = CommitDraft(0, "EVAL", "Synthetic group member", "Member " & i, #1/1/2026#, #9/30/2026#, "", "Test")
        stage = "Read member": ids(i) = id: Set f = SavedFields(id)
        f("RankRate") = "IT2": f("PromotionStatus") = "REGULAR": f("ComponentActive") = "Yes"
        f("ReportRegular") = "Yes": f("ReportingSeniorName") = "SYNTHETIC SENIOR": f("ReportingSeniorSSN") = "000000001": f("UnitUIC") = "12345"
        f("ObservationStatus") = "Observed": f("PromotionRecommendation") = "Promotable"
        If i = 1 Or i = 3 Then f("PromotionRecommendation") = "Early Promote"
        If i = 2 Then f("ComponentActive") = "No": f("ComponentTAR") = "Yes"
        If i = 3 Then f("ObservationStatus") = "NOB": f("RankRate") = "IT1": f("PeriodEnd") = #8/31/2026#
        stage = "Commit member": Call SaveFixture(id, f, "Synthetic group keys")
        stage = "Add member": AddGroupMember g, id, "E5"
    Next
    text = CapacityText(g)
    Check InStr(text, "2 observed; EP 1/1") > 0, "NOB contaminated count or ACT/TAR split"
    rev = CurrentRevision(ids(1)): operations = DCount("*", "OperationJournal")
    text = CapacityText(g)
    Check CurrentRevision(ids(1)) = rev And DCount("*", "OperationJournal") = operations, "Diagnostics mutated saved state"
    Set f = SavedFields(ids(2)): f("PromotionRecommendation") = "Early Promote"
    Call SaveFixture(ids(2), f, "Exceeded EP")
    text = CapacityText(g): Check InStr(text, "exceeds") > 0 And InStr(text, "EP 2/1") > 0, "Over-capacity observed roster not identified"
    f("PromotionStatus") = "FROCKED"
    Call SaveFixture(ids(2), f, "Mixed promotion status")
    Check InStr(CapacityText(g), "Capacity not calculated") > 0, "Mixed keys produced capacity"
    f("PromotionStatus") = "REGULAR": f("ReportingSeniorSSN") = ""
    Call SaveFixture(ids(2), f, "Missing senior")
    Check InStr(CapacityText(g), "Capacity not calculated") > 0, "Missing identity produced capacity"
    f("ReportingSeniorSSN") = "000000001": f("ComponentActive") = "Yes"
    Check Not GroupKeysResolved(f, "E5", "EVAL"), "Multiple component selections accepted"
    f("ComponentActive") = "No": f("RankRate") = "IT1"
    Check Not GroupKeysResolved(f, "E5", "EVAL"), "Wrong worn grade accepted"
    f("RankRate") = "IT2": f("FormKind") = "FITREP": f("RankRate") = "LT": f("Designator") = "6110"
    Check Not GroupKeysResolved(f, "O3", "FITREP"), "Unresolved competitive category accepted"
    Check InStr(ValidationPolicyReference("CAPACITY"), "1-17") > 0, "Capacity policy reference missing"
    stage = "category basis"
    f("Designator") = "6110": Check VCategory(f, "Active") = "Active LDO 61", "Active LDO category missing"
    f("Designator") = "6315": Check VCategory(f, "Reserve") = "Reserve LDO line", "Reserve LDO combined line missing"
    f("Designator") = "6515": Check VCategory(f, "Reserve") = "Reserve LDO staff", "Reserve staff split missing"
    f("Designator") = "7401": Check VCategory(f, "Active") = "Active CWO 72", "740x special category missing"
    f("Designator") = "7371": Check VCategory(f, "Active") = "Active CWO UAV", "UAV exception missing"
    Check VCategory(f, "Reserve") = "Reserve CWO", "Reserve CWO combined category missing"
    Check VCategory(f, "Unconfirmed") = "", "Unconfirmed category was inferred from duty status"
    stage = "UIC review"
    Set f = SavedFields(ids(2)): f("ReportingSeniorSSN") = "000000001": f("PromotionStatus") = "REGULAR": f("UnitUIC") = "67890": f("Trait33") = "4.0"
    SaveFixture ids(2), f, "Mixed enlisted UIC"
    Check InStr(CapacityText(g), "Capacity not calculated") > 0, "Unconfirmed multi-UIC roster accepted"
    stamp = GroupStamp(g): rev = CurrentRevision(ids(2)): snapshots = DCount("*", "GroupSnapshot")
    SaveGroupPolicy g, stamp, "Combined enlisted UICs", "Unconfirmed", "Synthetic reporting senior reviewed combined UIC roster", True
    Check GroupReviewCurrent(g), "Recorded review immediately stale"
    Check CurrentRevision(ids(2)) = rev And DCount("*", "GroupSnapshot") = snapshots, "Policy review rewrote report or froze roster"
    Check InStr(CapacityText(g), "EP 2/1") > 0, "Confirmed UIC exception still withheld capacity"
    Check InStr(GroupImpactText(g), "arithmetic average: 4.00") > 0, "Group consequence average missing"
    before = GroupPolicyValue(g, "Note")
    On Error Resume Next
    SaveGroupPolicy g, stamp, "Combined enlisted UICs", "Unconfirmed", "Stale must reject", True
    i = Err.Number: Err.Clear: On Error GoTo Failed
    Check i <> 0 And GroupPolicyValue(g, "Note") = before, "Stale review overwrote context"
    Set overrides = CreateObject("Scripting.Dictionary"): Set f = SavedFields(ids(2)): f("PeriodEnd") = #8/31/2026#: overrides.Add CStr(ids(2)), f
    Check InStr(GroupImpactText(g, overrides), "Capacity not calculated") > 0, "Proposed grouping changes not assessed"
    Check InStr(GroupImpactText(g, overrides), "fresh roster review") > 0, "Proposal hid review consequence"
    Check CurrentRevision(ids(2)) = rev And GroupReviewCurrent(g), "Read-only preview changed review baseline"
    SaveFixture ids(2), f, "Changed group key"
    Check Not GroupReviewCurrent(g) And InStr(GroupReviewLabel(g), "stale") > 0, "Saved change did not stale completeness"
    f("PeriodEnd") = #9/30/2026#: f("ReportingSeniorUIC") = "12345"
    SaveFixture ids(2), f, "Restore keys"
    SaveGroupPolicy g, GroupStamp(g), "Separate enlisted UICs", "Unconfirmed", "Synthetic separate primary-UIC review", False
    Check InStr(CapacityText(g), "Capacity not calculated") > 0, "Separated member/primary UIC mismatch accepted"
    stage = "native policy review"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
    Check UXOpenGroupPolicy(), "Group policy entry failed"
    Check Forms!frmUXGroupPolicy!chkComplete = False, "Completeness was checked automatically"
    Check InStr(Forms!frmUXGroupPolicy!txtEvidence, "Table 1-4") > 0, "Policy locator absent from review"
    DoCmd.Close acForm, "frmUXGroupPolicy", acSaveNo
    Check UXChooseMembers(), "Membership entry failed"
    Check Not UXApplyGroupMembers(), "Unpreviewed membership applied"
    Check UXPreviewGroupMembers(), "Native membership consequence preview failed"
    Check InStr(Forms!frmUXMembers!txtImpact, "AFTER SELECTED MEMBERS") > 0, "Membership before/after missing"
    Forms!frmUXMembers!txtNote = "Changed after preview"
    Check Not UXApplyGroupMembers(), "Changed note bypassed preview guard"
    SaveGroupPolicy g, GroupStamp(g), "Separate enlisted UICs", "Unconfirmed", "Synthetic complete roster before membership change", True
    For i = 0 To Forms!frmUXMembers!lstReports.ListCount - 1
        If Nz(Forms!frmUXMembers!lstReports.Column(0, i), 0) = ids(3) Then Forms!frmUXMembers!lstReports.Selected(i) = False
    Next
    Check UXPreviewGroupMembers(), "Changed membership could not be previewed again"
    Check UXApplyGroupMembers(), "Fresh membership preview could not be applied"
    Check DCount("*", "GroupMember", "GroupID=" & g & " AND ReportID=" & ids(3)) = 0 And Not GroupReviewCurrent(g), "Applied membership did not invalidate completeness"
    Check InStr(GroupImpactText(g, , ""), "promotion capacity is unavailable") > 0, "Empty proposed roster implied capacity"
    TestGroupCorrectness = "PASS: " & checks & " group correctness assertions": Exit Function
Failed:
    TestGroupCorrectness = "FAIL after " & checks & " at " & stage & ": " & Err.Number & " " & Err.Description
End Function
