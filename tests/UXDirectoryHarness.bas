Attribute VB_Name = "UXDirectoryHarness"
Option Explicit
Private n As Long
Private stage As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6981, , message
    n = n + 1
End Sub
Public Function TestUXDirectory() As String
    Dim m As Long, f As Access.Form, b As Long, id As Long, reportID As Long, rev As Long, other As Long
    On Error GoTo Failed
    ConfigureWorkspace "DIRECTORY UI SYNTHETIC", "Directory tester"
    DoCmd.OpenForm "frmUXDirectory"
    UXDirectoryEdit True: Set f = Forms!frmUXMemberEditor
    f!txtName = "Casey Morgan": f!txtReportName = "MORGAN, CASEY": f!txtRank = "IT2": f!txtSSN = "000000005"
    Check UXDirectorySave(), "Member UI save failed"
    m = DLookup("MemberID", "CommandMember", "MemberSSN='000000005'")
    Check Forms!frmUXDirectory!lstMembers.ListCount >= 2, "Member list missing row/header"
    Forms!frmUXDirectory!lstMembers = m
    Check UXDirectoryEdit(False), "Existing profile did not open"
    Check Forms!frmUXMemberEditor!txtReportName = "MORGAN, CASEY", "Separate report name not retained"
    Forms!frmUXMemberEditor!txtSSN = "bad"
    Check Not UXDirectorySave(), "Invalid SSN UI save accepted"
    Check Forms!frmUXMemberEditor!txtSSN = "bad" And InStr(Forms!frmUXMemberEditor!lblStatus.Caption, "nine") > 0, "Error cleared pending identity input"
    UXDirectoryClose "frmUXMemberEditor"
    Check UXDirectoryHistory(), "History did not open"
    Check UXPriorOpen(True), "Prior editor did not open"
    Set f = Forms!frmUXPriorEditor
    f!cboKind = "CHIEFEVAL": UXPriorLabels
    Check f!lblTrait33.Caption = UXFieldCaption("Trait33", "CHIEFEVAL"), "Chief trait caption wrong"
    f!txtFrom = #1/1/2025#: f!txtTo = #9/30/2025#: f!cboPromotion = "Promotable": f!txtSource = "Synthetic check-in evaluation"
    For b = 33 To 39: f.Controls("cboTrait" & b) = "4.0": Next
    Check UXPriorSave(), "Prior UI save failed"
    id = DLookup("PriorID", "PriorEvaluation", "MemberID=" & m)
    Check PriorEvaluationITA(id) = "4.00", "Prior ITA incorrect"
    Forms!frmUXMemberHistory!lstHistory = id
    Check UXPriorOpen(False), "Prior edit did not open"
    Check Forms!frmUXPriorEditor!txtSource = "Synthetic check-in evaluation" And Forms!frmUXPriorEditor!cboTrait39 = "4.0", "Prior reopen lost context/trait"
    UXDirectoryClose "frmUXPriorEditor"
    reportID = CommitDraft(0, "EVAL", "History identity", "Casey Morgan", #1/1/2026#, #9/30/2026#, "", "Test", "")
    LinkReportDirectoryMember reportID, m, CurrentRevision(reportID), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & m)), "Synthetic identity"
    Call UXHistoryLoad: Forms!frmUXMemberHistory!lstHistory = -reportID
    Check UXHistoryConfirmIdentity(), "Identity preview absent"
    Check Not UXHistoryApplyIdentity(), "Reasonless confirmation accepted"
    Forms!frmUXConfirmIdentity!txtReason = "Synthetic checked SSN"
    rev = CurrentRevision(reportID)
    Check UXHistoryApplyIdentity() And CurrentRevision(reportID) = rev, "Identity confirmation failed or revised report"
    other = SaveDirectoryMember(0, 0, "Jordan Reed", "REED, JORDAN", "ITC", "000000006")
    Forms!frmUXMemberHistory!lstHistory = -reportID: UXHistoryConfirmIdentity
    Set f = Forms!frmUXConfirmIdentity: f!cboMember = other: f!txtReason = "Synthetic corrected identity"
    Check Not UXHistoryApplyIdentity(), "Changed member selection applied without preview"
    Check UXHistoryIdentityPreview() And InStr(f!txtPreview, "Jordan Reed") > 0, "Chosen member not previewed"
    Call SaveDirectoryMember(other, 1, "Jordan Reed updated", "REED, JORDAN", "ITC", "000000006")
    Check Not UXHistoryApplyIdentity(), "Stale selected member accepted"
    Check UXHistoryIdentityPreview(), "Fresh identity preview failed"
    Check UXHistoryApplyIdentity() And DLookup("MemberID", "MemberReport", "ReportID=" & reportID) = other And CurrentRevision(reportID) = rev, "Corrected member link failed or rewrote report"
    UXDirectoryClose "frmUXMemberHistory"
    ' An edited profile updates the member's open (draft) reports and names the rest.
    Dim p As Long, draft As Long, signed As Long, reviewed As Long, reviewer As Long, untouched As Long, linked As Variant
    stage = "profile fixture"
    p = SaveDirectoryMember(0, 0, "Riley Park", "PARK, RILEY", "IT2", "000000007")
    draft = CommitDraft(0, "EVAL", "Open draft", "Riley Park", #1/1/2026#, #9/30/2026#, "", "Test", "")
    signed = CommitDraft(0, "EVAL", "Signed report", "Riley Park", #1/1/2025#, #9/30/2025#, "", "Test", "")
    reviewed = CommitDraft(0, "EVAL", "Report under review", "Riley Park", #10/1/2025#, #12/31/2025#, "", "Test", "")
    For Each linked In Array(draft, signed, reviewed)
        LinkReportDirectoryMember CLng(linked), p, CurrentRevision(CLng(linked)), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & p)), "Synthetic link"
    Next
    Call MarkReportSigned(signed, CurrentRevision(signed), #10/5/2025#)
    reviewer = AddContributor("Profile reviewer")
    Call AssignReview(reviewed, CurrentRevision(reviewed), reviewer, Null, "Check it")
    untouched = CurrentRevision(signed)
    stage = "open profile"
    Forms!frmUXDirectory!lstMembers.Requery: Forms!frmUXDirectory!lstMembers = p
    Check UXDirectoryEdit(False), "Profile for report update did not open"
    Forms!frmUXMemberEditor!txtRank = "IT1"
    stage = "save profile"
    Check UXDirectorySave(), "Profile save with open reports failed: " & Forms!frmUXMemberEditor!lblStatus.Caption
    Check CurrentProject.AllForms("frmUXMemberEditor").IsLoaded, "Editor closed without reporting the report update"
    Check InStr(Forms!frmUXMemberEditor!lblStatus.Caption, "Updated 1 open report") > 0, "Update count not shown: " & Forms!frmUXMemberEditor!lblStatus.Caption
    Check InStr(Forms!frmUXMemberEditor!lblStatus.Caption, "report " & signed & " (signed)") > 0 And InStr(Forms!frmUXMemberEditor!lblStatus.Caption, "report " & reviewed & " (under review)") > 0, "Skipped reports not named: " & Forms!frmUXMemberEditor!lblStatus.Caption
    Check DLookup("RankRate", "Report", "ReportID=" & draft) = "IT1", "Open draft did not take the profile change"
    stage = "skipped reports"
    Check IsNull(DLookup("RankRate", "Report", "ReportID=" & signed)) And CurrentRevision(signed) = untouched, "Signed report changed"
    Check IsNull(DLookup("RankRate", "Report", "ReportID=" & reviewed)), "Report under review changed"
    Forms!frmUXMemberEditor!txtRank = "IT1": Forms!frmUXMemberEditor!txtReportName = "PARK, RILEY J"
    Check UXDirectorySave(), "Second save from the open editor failed (stale version)"
    Check DLookup("OfficialName", "Report", "ReportID=" & draft) = "PARK, RILEY J", "Second profile change not applied"
    UXDirectoryClose "frmUXMemberEditor": UXDirectoryClose "frmUXDirectory"
    TestUXDirectory = "PASS: " & n & " native directory/history UI assertions": Exit Function
Failed:
    TestUXDirectory = "FAIL at " & stage & " after " & n & ": " & Err.Number & " " & Err.Description
End Function
