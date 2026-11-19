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
    ' Continuity: every period in date order across report types; a gap is only possible,
    ' an overlap only counts between Swift Regular reports, and a repeated period counts once.
    Dim c As Long, marks As Object, rows As Collection, s1 As Long, s2 As Long, s3 As Long, conc As Long, i As Long, gapRow As Long
    stage = "continuity fixture"
    c = SaveDirectoryMember(0, 0, "Alex Rowe", "ROWE, ALEX", "IT1", "000000008")
    Set marks = CreateObject("Scripting.Dictionary")
    For b = 33 To 39: marks.Add "Trait" & b, "4.0": Next
    Call SavePriorEvaluation(0, 0, c, "EVAL", "IT1", #11/16/2023#, #11/15/2024#, "Promotable", "Observed", marks, "Synthetic prior")
    s1 = CommitDraft(0, "EVAL", "Continuity one", "Alex Rowe", #1/1/2025#, #11/15/2025#, "", "Test", "")
    s2 = CommitDraft(0, "CHIEFEVAL", "Continuity two", "Alex Rowe", #11/16/2025#, #11/15/2026#, "", "Test", "")
    For Each linked In Array(s1, s2)
        LinkReportDirectoryMember CLng(linked), c, CurrentRevision(CLng(linked)), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & c)), "Synthetic link"
    Next
    CurrentDb.Execute "UPDATE Report SET ReportRegular='Yes' WHERE ReportID IN (" & s1 & "," & s2 & ")", dbFailOnError
    'The same report also recorded as a prior, and an overlapping prior (possibly Concurrent).
    Call SavePriorEvaluation(0, 0, c, "CHIEFEVAL", "ITC", #11/16/2025#, #11/15/2026#, "Promotable", "Observed", marks, "Synthetic copy")
    Call SavePriorEvaluation(0, 0, c, "EVAL", "IT1", #6/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Synthetic overlap")
    Set rows = MemberContinuityRows(c)
    Check rows.Count = 1, "Continuity rows: expected one gap, got " & rows.Count
    Check rows(1)(0) = "Gap" And rows(1)(1) = #11/16/2024# And rows(1)(2) = #12/31/2024# And rows(1)(3) = 46, "Gap dates wrong"
    Check MemberContinuitySummary(rows) = "Continuity: 1 possible gap. Prior reports do not record Regular or Concurrent; verify on BOL Continuity.", "Gap summary wrong: " & MemberContinuitySummary(rows)
    stage = "continuity overlap"
    s3 = CommitDraft(0, "CHIEFEVAL", "Continuity three", "Alex Rowe", #10/1/2026#, #9/30/2027#, "", "Test", "")
    conc = CommitDraft(0, "CHIEFEVAL", "Concurrent only", "Alex Rowe", #1/1/2029#, #6/30/2029#, "", "Test", "")
    For Each linked In Array(s3, conc)
        LinkReportDirectoryMember CLng(linked), c, CurrentRevision(CLng(linked)), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & c)), "Synthetic link"
    Next
    CurrentDb.Execute "UPDATE Report SET ReportRegular='Yes' WHERE ReportID=" & s3, dbFailOnError
    CurrentDb.Execute "UPDATE Report SET ReportRegular='No',ReportConcurrent='Yes' WHERE ReportID=" & conc, dbFailOnError
    Set rows = MemberContinuityRows(c)
    Check rows.Count = 2, "Expected a gap and an overlap; Concurrent report must not add a gap. Got " & rows.Count
    Check rows(2)(0) = "Overlap" And rows(2)(1) = #10/1/2026# And rows(2)(2) = #11/15/2026#, "Overlap dates wrong"
    Check InStr(MemberContinuitySummary(rows), "1 possible gap, 1 overlap between Regular reports.") > 0, "Overlap summary wrong: " & MemberContinuitySummary(rows)
    stage = "continuity history list"
    Check UXDirectoryHistoryFor(c), "History did not open for continuity"
    Set f = Forms!frmUXMemberHistory
    gapRow = -1
    For i = 0 To f!lstHistory.ListCount - 1
        If f!lstHistory.Column(1, i) = "Possible gap" Then gapRow = i
    Next
    Check gapRow > 0 And gapRow < f!lstHistory.ListCount - 1, "Gap row missing or out of place"
    Check f!lstHistory.Column(0, gapRow) = "0" And f!lstHistory.Column(3, gapRow) = "2024-11-16 to 2024-12-31" And f!lstHistory.Column(5, gapRow) = "46 days", "Gap row content wrong"
    Check f!lstHistory.Column(3, gapRow + 1) = "2023-11-16 to 2024-11-15" And f!lstHistory.Column(3, gapRow - 1) = "2025-01-01 to 2025-11-15", "Gap row not between its reports"
    Check f!lblStatus.Caption = MemberContinuitySummary(rows), "Continuity summary not shown"
    f!lstHistory = 0
    Check Not UXPriorOpen(False) And Not CurrentProject.AllForms("frmUXPriorEditor").IsLoaded And f!lblStatus.Caption = "Choose a prior evaluation to edit.", "A continuity row opened the prior editor"
    UXDirectoryClose "frmUXMemberHistory"
    TestUXDirectory = "PASS: " & n & " native directory/history UI assertions": Exit Function
Failed:
    TestUXDirectory = "FAIL at " & stage & " after " & n & ": " & Err.Number & " " & Err.Description
End Function
