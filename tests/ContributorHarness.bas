Attribute VB_Name = "ContributorHarness"
Option Explicit
'P06: contributors from the directory, deleting contributors, people fields from members.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1890, , message
    checks = checks + 1
End Sub
Private Function Refused(ByVal action As String, ByVal id As Long) As Boolean
    On Error GoTo Failed
    Select Case action
        Case "delete": Call DeleteContributor(id)
        Case "member": Call AddContributorFromMember(id)
        Case "switch": Call SwitchContributor(id)
        Case "counselor": Call MemberPeopleText(id, "CounselorName")
    End Select
    Exit Function
Failed:
    Refused = True
End Function
Private Function Steps(ByVal run As Long) As Long
    Steps = DCount("*", "RouteRunStep", "RunID=" & run)
End Function
Public Function TestContributors() As String
    Dim requester As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, chain As Long, solo As Long, pair As Long
    Dim id1 As Long, id2 As Long, id3 As Long, id4 As Long, run1 As Long, run2 As Long, run4 As Long, member As Long, other As Long, profile As Long, summary As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC CONTRIBUTORS", "Requester"
    requester = MajorAuthor()

    ' A contributor from the directory takes rank and report name, once per member.
    member = SaveDirectoryMember(0, 0, "JONES PAUL ALLEN", "JONES, PAUL A", "LCDR", "")
    profile = AddContributorFromMember(member)
    Check DLookup("DisplayName", "UserProfile", "UserProfileID=" & profile) = "LCDR JONES, PAUL A" And DLookup("MemberID", "UserProfile", "UserProfileID=" & profile) = member, "Directory contributor name or link wrong"
    Check Refused("member", member), "The same member was added twice"
    Check AddContributor("Manual Person") > 0, "Manual contributor not added"

    ' People fields from a member.
    Check MemberPeopleText(member, "RaterName") = "JONES, P A, LCDR" And MemberPeopleText(member, "SeniorRaterName") = "JONES, P A, LCDR" And MemberPeopleText(member, "CounselorName") = "JONES, P A", "Rater or counselor text wrong"
    Check MemberPeopleText(member, "CounselorLastName") = "JONES" And MemberPeopleText(member, "CounselorFirstInitial") = "P" And MemberPeopleText(member, "CounselorMiddleInitial") = "A", "Counselor name parts wrong"
    other = SaveDirectoryMember(0, 0, "SYNTHETICLASTNAMEX ANNE BETH", "SYNTHETICLASTNAMEX, ANNE B", "CAPT", "")
    Check MemberPeopleText(other, "RaterName") = "SYNTHETICLASTNAMEX, CAPT" And MemberPeopleText(other, "CounselorName") = "SYNTHETICLASTNAMEX", "Long name not shortened to the last name"
    other = SaveDirectoryMember(0, 0, "ANNE SYNTHETICVERYLONGLASTNAMEZ", "", "", "")
    Check Refused("counselor", other), "A name that cannot fit was accepted"

    ' Routes and a review request involving the contributors.
    a = AddContributor("Division Officer"): b = AddContributor("Department Head"): c = AddContributor("Executive Officer")
    d = AddContributor("Solo Reviewer"): e = AddContributor("First Reader"): f = AddContributor("Final Reader")
    chain = CreateRouteTemplate("Chain", a & "," & b & "," & c, "Review")
    solo = CreateRouteTemplate("Solo", CStr(d), "Review")
    pair = CreateRouteTemplate("Pair", e & "," & f, "Review")
    id1 = CommitDraft(0, "EVAL", "Chain report", "SYNTHETIC ONE", #1/1/2026#, #9/30/2026#, "")
    id2 = CommitDraft(0, "EVAL", "Solo report", "SYNTHETIC TWO", #1/1/2026#, #9/30/2026#, "")
    id3 = CommitDraft(0, "EVAL", "Assigned report", "SYNTHETIC THREE", #1/1/2026#, #9/30/2026#, "")
    id4 = CommitDraft(0, "EVAL", "Pair report", "SYNTHETIC FOUR", #1/1/2026#, #9/30/2026#, "")
    run1 = StartRoute(id1, chain): run2 = StartRoute(id2, solo): run4 = StartRoute(id4, pair)
    Call AssignReview(id3, CurrentRevision(id3), b, Null, "Please review")
    Check SwitchContributor(a), "Switch to step 1": ForwardRoute run1, ""
    Check SwitchContributor(e), "Switch to first reader": ForwardRoute run4, ""
    Check SwitchContributor(requester), "Switch back to the requester"
    Check ReviewHolder(id1) = b And ReviewHolder(id4) = f, "Setup holders wrong"

    ' The active contributor cannot delete themself.
    Check Refused("delete", requester), "The active contributor was deleted"

    ' Deleting the holder: the next step holds the draft; the template closes the gap.
    summary = DeleteContributor(b)
    Check ReviewHolder(id1) = c And Steps(run1) = 2 And DLookup("CurrentStep", "RouteRun", "RunID=" & run1) = 2, "Next step does not hold the report after the holder was deleted"
    Check DCount("*", "RouteRunEvent", "RunID=" & run1 & " AND EventKind='Reviewer removed'") = 1, "Removal not recorded on the route"
    Check DCount("*", "RouteTemplateStep", "TemplateID=" & chain) = 2 And DLookup("AssigneeID", "RouteTemplateStep", "TemplateID=" & chain & " AND StepOrdinal=2") = c, "Template steps not closed up"
    Check DLookup("TaskState", "ReviewTask", "ReportID=" & id3) = "Cancelled", "Review request not cancelled"
    Check Refused("switch", b) And DCount("*", "UserProfile", "IsRetired=False AND UserProfileID=" & b) = 0, "Deleted contributor still usable"
    Check DLookup("DisplayName", "UserProfile", "UserProfileID=" & b) = "Department Head" And InStr(summary, "1 review request(s) cancelled") > 0, "Name or summary wrong: " & summary

    ' Deleting an earlier step keeps the same holder at a lower step number.
    Call DeleteContributor(a)
    Check ReviewHolder(id1) = c And Steps(run1) = 1 And DLookup("CurrentStep", "RouteRun", "RunID=" & run1) = 1, "Earlier step removal moved the holder"

    ' The only reviewer: template deleted and route cancelled.
    Call DeleteContributor(d)
    Check DLookup("IsArchived", "RouteTemplate", "TemplateID=" & solo) = True And DLookup("RunState", "RouteRun", "RunID=" & run2) = "Cancelled", "Empty template or route left active"

    ' The removed holder was the last step: the route is complete with the last reviewer left.
    Call DeleteContributor(f)
    Check DLookup("RunState", "RouteRun", "RunID=" & run4) = "Accepted" And ReviewHolder(id4) = e, "Route not completed after the last step was deleted"

    ' Adding the same name again restores the old profile.
    Check AddContributor("Department Head") = b And DCount("*", "UserProfile", "IsRetired=False AND UserProfileID=" & b) = 1, "Re-added contributor not restored"

    ' The editor's people fields open the picker and take the chosen member.
    id1 = CommitDraft(0, "EVAL", "Picker report", "SYNTHETIC FIVE", #1/1/2026#, #9/30/2026#, "")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id1)
    Check UXPickPerson("RaterName"), "Rater picker did not open"
    Forms!frmUXSeniorPicker!lstMembers = member: Check RSPickerUse(), "Use selected member failed for the rater"
    Check Forms!frmEditor!txtRaterName = "JONES, P A, LCDR" And Not CurrentProject.AllForms("frmUXSeniorPicker").IsLoaded, "Rater field not filled from the member"
    Check UXPickPerson("CounselorName"), "Counselor picker did not open"
    Forms!frmUXSeniorPicker!lstMembers = member: Check RSPickerUse(), "Use selected member failed for the counselor"
    Check Forms!frmEditor!txtCounselorName = "JONES, P A" And EditorHasPending(), "Counselor field not filled or not pending"
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
    TestContributors = "PASS: " & checks & " contributor directory, delete and people-field assertions": Exit Function
Failed:
    TestContributors = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
