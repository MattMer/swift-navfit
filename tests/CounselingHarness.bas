Attribute VB_Name = "CounselingHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 9590, , message
    checks = checks + 1
End Sub
Private Function Rejected(ByVal memberID As Long, ByVal counseledOn As Date, ByVal outcome As String, ByVal counselor As String, ByVal reason As String, Optional ByVal disposition As String = "", Optional ByVal dispositionOn As Variant = Null, Optional ByVal id As Long = 0, Optional ByVal version As Long = 0) As Boolean
    On Error GoTo Failed
    Call SaveCounseling(id, version, memberID, counseledOn, outcome, counselor, reason, disposition, dispositionOn)
    Exit Function
Failed:
    Rejected = True
End Function
Private Function MonthsText(ByVal grade As String, ByVal fromDate As Date, ByVal toDate As Date) As String
    Dim due As Collection, i As Long
    Set due = CounselingDueMonths(grade, fromDate, toDate)
    For i = 1 To due.Count
        If i > 1 Then MonthsText = MonthsText & ","
        MonthsText = MonthsText & Format$(due(i), "yyyy-mm")
    Next
End Function
Private Function Member(ByVal name As String, ByVal rank As String) As Long
    Member = SaveDirectoryMember(0, 0, name, UCase$(name), rank, "")
End Function
Private Function RegularReport(ByVal memberID As Long, ByVal fromDate As Date, ByVal toDate As Date, Optional ByVal regular As String = "Yes") As Long
    Dim id As Long, changes As Object
    id = CommitDraft(0, "EVAL", "Counseling carry-over", "Synthetic member", fromDate, toDate, "", "Test", "")
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "ReportRegular", regular
    If regular <> "Yes" Then changes.Add "ReportConcurrent", "Yes"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Report type")
    LinkReportDirectoryMember id, memberID, CurrentRevision(id), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID)), "Synthetic identity"
    RegularReport = id
End Function
Private Function Suggested(ByVal reportID As Long) As String
    Dim s As Object
    Set s = CounselingSuggestion(SavedFields(reportID), reportID)
    If s Is Nothing Then Suggested = "(none)" Else Suggested = s("CounselingDate") & "|" & Nz(s("CounselorName"), "")
End Function
Private Function HasFinding(ByVal reportID As Long) As Boolean
    Dim item As Object
    For Each item In ValidateLiveFields(SavedFields(reportID), reportID)
        If item("Code") = "COUNSEL-RECORD" Then HasFinding = True
    Next
End Function
Public Function TestCounseling() As String
    Dim e5 As Long, blank As Long, chief As Long, id As Long, other As Long, report As Long, second As Long, concurrent As Long
    Dim restored As Long, stale As Long, rejectedRestore As Boolean
    Dim s As Object, changes As Object, f As Access.Form, i As Long, found As Boolean, before As Long, stamp As String, merge As Long, source As Long, target As Long, ops As Long, stage As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC COUNSELING", "Counseling tester"

    stage = "schedule"
    Check MonthsText("O6", #1/1/2026#, #12/31/2027#) = "2026-01,2026-11,2027-09", "Captain schedule wrong: " & MonthsText("O6", #1/1/2026#, #12/31/2027#)
    Check MonthsText("O5", #1/1/2026#, #12/31/2027#) = "2026-10,2027-12", "Commander schedule wrong"
    Check MonthsText("O4", #1/1/2026#, #12/31/2027#) = "2026-04,2027-04", "LCDR schedule wrong"
    Check MonthsText("O3", #1/1/2026#, #12/31/2026#) = "2026-07" And MonthsText("O2", #1/1/2026#, #12/31/2026#) = "2026-08" And MonthsText("O1", #1/1/2026#, #12/31/2026#) = "2026-11", "Junior officer schedule wrong"
    Check MonthsText("W4", #1/1/2026#, #12/31/2027#) = "2026-09,2027-09", "CWO3-5 schedule wrong"
    Check MonthsText("W2", #1/1/2026#, #12/31/2027#) = "2027-03" And MonthsText("W1", #1/1/2026#, #12/31/2026#) = "", "CWO2/WO1 schedule must start in 2027"
    Check MonthsText("E9", #1/1/2026#, #12/31/2027#) = "2026-10,2027-10", "E-9 schedule wrong"
    Check MonthsText("E7", #1/1/2026#, #12/31/2026#) = "2026-03" And MonthsText("E8", #1/1/2026#, #12/31/2026#) = "2026-03", "E-7/E-8 schedule wrong"
    Check MonthsText("E6", #1/1/2026#, #12/31/2026#) = "2026-05" And MonthsText("E5", #1/1/2026#, #12/31/2026#) = "2026-09" And MonthsText("E4", #1/1/2026#, #12/31/2026#) = "2026-12", "E-4 to E-6 schedule wrong"
    Check MonthsText("E1", #1/1/2026#, #12/31/2026#) = "2026-01" And MonthsText("E3", #1/1/2027#, #12/31/2027#) = "2027-01", "E-1 to E-3 schedule wrong"
    Check MonthsText("E5", #10/1/2026#, #8/31/2027#) = "", "Range filter wrong"

    stage = "status"
    e5 = Member("Synthetic Counseled", "IT2")
    Check CounselingMemberStatus(e5, #6/15/2026#)("Status") = "Scheduled", "Far due month not Scheduled"
    Check CounselingMemberStatus(e5, #8/15/2026#)("Status") = "Due soon", "Due within 60 days not Due soon"
    Check CounselingMemberStatus(e5, #9/10/2026#)("Status") = "Due this month", "Due month not detected"
    Check CounselingMemberStatus(e5, #10/5/2026#)("Status") = "Overdue", "Missed month not Overdue"
    Check Format$(CounselingMemberStatus(e5, #10/5/2026#)("DueMonth"), "yyyy-mm") = "2026-09", "Overdue month wrong"
    blank = Member("Synthetic No Rank", "")
    Check CounselingMemberStatus(blank, #10/5/2026#)("Status") = "Enter rank/rate", "Missing rank not reported"
    chief = Member("Synthetic Master Chief", "ITCM")
    Check CounselingMemberStatus(chief, #10/5/2026#)("Status") = "Not scheduled" And InStr(CounselingMemberStatus(chief, #10/5/2026#)("Note"), "No Regular periodic report") = 1, "E-9 still shown as due after NAVADMIN 132/26"
    Check CounselingMemberStatus(chief, #2/10/2026#)("Status") <> "Not scheduled", "E-9 exempted before 3 Jun 2026"
    Check CounselingMemberStatus(Member("Synthetic CWO5", "CWO5"), #10/5/2026#)("Status") = "Not scheduled", "CWO5 still shown as due"

    stage = "save rules"
    Check Rejected(e5, #9/20/2026#, "Done", "EXAMPLE, C", ""), "Unknown outcome accepted"
    Check Rejected(e5, #9/20/2026#, "Completed", "", ""), "Completed without counselor accepted"
    Check Rejected(e5, #9/20/2026#, "Completed", String$(21, "C"), ""), "Overlong counselor accepted"
    Check Rejected(e5, #9/20/2026#, "Not performed", "", ""), "Not performed without reason accepted"
    Check Rejected(e5, Date + 1, "Completed", "EXAMPLE, C", ""), "Future date accepted"
    Check Rejected(e5, #9/20/2026#, "Completed", "EXAMPLE, C", "", "Shredded", Date), "Unknown disposition accepted"
    Check DCount("*", "MemberCounseling") = 0, "Rejected save left a row"
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO MemberCounseling (CounselingGuid,MemberID,CounseledOn,Outcome,VersionCounter,CreatedBy,ModifiedBy) VALUES ('{X}'," & e5 & ",#2026-09-20#,'Completed',1,1,1)", dbFailOnError
    Check Err.Number <> 0, "Table rule allowed a completed row without counselor": Err.Clear
    On Error GoTo Failed

    stage = "save and edit"
    ops = DCount("*", "OperationJournal", "OperationKind='Save counseling record'")
    id = SaveCounseling(0, 0, e5, #9/20/2026#, "Completed", "EXAMPLE, C", "", "", Null)
    Check DLookup("VersionCounter", "MemberCounseling", "CounselingID=" & id) = 1 And DLookup("NotPerformedReason", "MemberCounseling", "CounselingID=" & id) & "" = "", "New record wrong"
    Check DCount("*", "OperationJournal", "OperationKind='Save counseling record'") = ops + 1, "Save was not journaled"
    Check CounselingMemberStatus(e5, #10/5/2026#)("Status") = "Scheduled", "Recorded session still overdue"
    Check Format$(CounselingMemberStatus(e5, #10/5/2026#)("LastCounseled"), "yyyy-mm-dd") = "2026-09-20", "Last counseled wrong"
    Check CounselingMemberStatus(e5, #9/25/2026#)("Status") = "Recorded", "In-month session not Recorded"
    Check Rejected(e5, #9/21/2026#, "Completed", "EXAMPLE, D", "", "", Null, id, 7), "Stale version accepted"
    Check SaveCounseling(id, 1, e5, #9/21/2026#, "Completed", "EXAMPLE, D", "", "", Null) = id, "Edit failed"
    Check DLookup("VersionCounter", "MemberCounseling", "CounselingID=" & id) = 2 And DLookup("CounselorName", "MemberCounseling", "CounselingID=" & id) = "EXAMPLE, D", "Edit not stored"
    other = SaveCounseling(0, 0, e5, #3/2/2026#, "Not performed", "ignored", "TEMADD", "", Null)
    Check IsNull(DLookup("CounselorName", "MemberCounseling", "CounselingID=" & other)) And DLookup("NotPerformedReason", "MemberCounseling", "CounselingID=" & other) = "TEMADD", "Not performed kept a counselor"
    Check ArchiveCounseling(other, 1), "Archive failed"
    Check Rejected(e5, #3/2/2026#, "Not performed", "", "ILLNESS", "", Null, other, 2), "Archived record edited"
    Check DLookup("IsArchived", "MemberCounseling", "CounselingID=" & other) = True, "Archive not stored"

    stage = "carry-over"
    report = RegularReport(e5, #9/16/2025#, #9/15/2026#)
    Check Suggested(report) = "NOT REQ|", "Session after the period not answered with NOT REQ: " & Suggested(report)
    Check RegularReport(blank, #9/16/2025#, #9/15/2026#) > 0, "Unrecorded member report failed"
    Check Suggested(DMax("ReportID", "Report")) = "(none)", "Member without records got a suggestion"
    report = RegularReport(e5, #9/16/2026#, #9/15/2027#)
    Check Suggested(report) = "2026-09-21|EXAMPLE, D", "Recorded session not suggested: " & Suggested(report)
    Check HasFinding(report), "Validation finding missing"
    second = RegularReport(e5, #11/1/2025#, #9/15/2026#)
    Check Suggested(second) = "NOT REQ|", "NOT REQ not suggested for empty block 30: " & Suggested(second)
    concurrent = RegularReport(e5, #9/16/2026#, #9/15/2027#, "No")
    Check Suggested(concurrent) = "(none)", "Concurrent report got a suggestion"
    ' A Concurrent-only report repeating the date does not use up the session for
    ' the Regular report.
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "CounselingDate", "2026-09-21": changes.Add "CounselorName", "EXAMPLE, D"
    CommitSavedFields concurrent, CurrentRevision(concurrent), changes, "Concurrent counseling entry"
    Check Suggested(report) = "2026-09-21|EXAMPLE, D", "Concurrent-only report hid the Regular suggestion: " & Suggested(report)

    stage = "report tools apply"
    before = CurrentRevision(report)
    DoCmd.OpenForm "frmEditor", , , , , , CStr(report)
    Check OpenDraftChecklist(report), "Validation view failed"
    Set f = UXToolsForm(): found = False
    For i = 0 To f!lstFindings.ListCount - 1
        If InStr(f!lstFindings.Column(1, i) & f!lstFindings.Column(3, i), "falls in this report period") > 0 Then f!lstFindings = f!lstFindings.ItemData(i): found = True
    Next
    Check found, "Counseling finding not listed"
    Call UXToolsFinding
    Check f!cmdApplyCounseling.Enabled, "Apply not enabled for counseling finding"
    Check UXToolsApplyCounseling(), "Apply failed"
    Check Forms!frmEditor!txtCounselingDate = "2026-09-21" And Forms!frmEditor!txtCounselorName = "EXAMPLE, D", "Apply did not fill blocks 30-31"
    Check EditorHasPending() And CurrentRevision(report) = before, "Apply saved or lost the pending edit"
    Check Suggested(report) = "2026-09-21|EXAMPLE, D", "Unsaved apply consumed the record"
    Check UXSaveEditor(), "Save after apply failed"
    Check Suggested(report) = "(none)" And Not HasFinding(report), "Saved report still offered the session"
    Check RegularReport(e5, #9/1/2026#, #9/15/2027#) > 0, "Overlapping report failed"
    ' The session is in use on another Regular report: no second offer, and no
    ' NOT REQ claiming nothing was recorded.
    Check Suggested(DMax("ReportID", "Report")) = "(none)", "Session in use offered again or turned into NOT REQ: " & Suggested(DMax("ReportID", "Report"))
    Call UXLeaveEditor("Discard")

    stage = "screens"
    Check UXDirectoryHistoryFor(e5), "History did not open"
    Check Forms!frmUXMemberHistory!lstCounseling.ListCount = 1 And Forms!frmUXMemberHistory!lstCounseling.Column(1, 0) = "2026-09-21", "History counseling list wrong"
    Check InStr(Forms!frmUXMemberHistory!lblCounselingStatus.Caption, "Counseling: ") = 1, "History counseling status missing"
    Check UXCounselingOpen(True), "Counseling editor did not open"
    Set f = Forms!frmUXCounselingEditor
    f!txtDate = #10/1/2026#: f!cboOutcome = "Not performed": UXCounselingOutcome
    Check Not f!txtCounselor.Enabled And f!cboReason.Enabled, "Outcome did not switch block 31 input"
    Check f!cboReason.ListCount >= 5 And Not f!cboReason.LimitToList And Not UXHasControl(f, "txtDispositionOn"), "Reason list or disposition date field wrong"
    f!cboReason = "FAMILY EMERGENCY"
    Check UXCounselingSave(), "Editor save failed"
    Check Forms!frmUXMemberHistory!lstCounseling.ListCount = 2, "History did not refresh"
    Forms!frmUXMemberHistory!lstCounseling = DMax("CounselingID", "MemberCounseling")
    Check UXCounselingOpen(False), "Edit did not open"
    Check Forms!frmUXCounselingEditor!cboReason = "FAMILY EMERGENCY", "Typed reason not kept"
    Forms!frmUXCounselingEditor!txtDate = Null
    Check Not UXCounselingSave() And Len(Forms!frmUXCounselingEditor!lblStatus.Caption) > 0, "Missing date accepted"
    UXDirectoryClose "frmUXCounselingEditor"
    Forms!frmUXMemberHistory!lstCounseling = DMax("CounselingID", "MemberCounseling")
    Check UXCounselingArchive(True) And Forms!frmUXMemberHistory!lstCounseling.ListCount = 1, "History archive failed"
    Forms!frmUXMemberHistory.Controls("chkArchivedCounseling") = True
    Check UXHistoryCounselingLoad() And Forms!frmUXMemberHistory!lstCounseling.ListCount = 2, "Archived counseling cannot be viewed"
    Check Not Forms!frmUXMemberHistory!cmdCounselingEdit.Enabled, "Archived counseling enabled editing"
    Forms!frmUXMemberHistory!lstCounseling = DMax("CounselingID", "MemberCounseling")
    restored = CLng(Forms!frmUXMemberHistory!lstCounseling): stale = CLng(DLookup("VersionCounter", "MemberCounseling", "CounselingID=" & restored))
    Check Not UXCounselingOpen(False), "Archived counseling opened for editing"
    Check UXCounselingArchive(True), "Archived counseling could not be restored"
    On Error Resume Next
    Call RestoreCounseling(restored, stale): rejectedRestore = (Err.Number = vbObjectError + 9507): Err.Clear
    On Error GoTo Failed
    Check rejectedRestore, "Stale counseling restore accepted"
    Check Not DLookup("IsArchived", "MemberCounseling", "CounselingID=" & DMax("CounselingID", "MemberCounseling")), "Restore did not persist"
    Forms!frmUXMemberHistory.Controls("chkArchivedCounseling") = False
    Check UXHistoryCounselingLoad() And Forms!frmUXMemberHistory!lstCounseling.ListCount = 2, "Restored counseling missing from active history"
    UXDirectoryClose "frmUXMemberHistory"
    DoCmd.OpenForm "frmMyWork"
    Check UXMyWorkMode("Counseling"), "Counseling tab failed"
    Check Forms!frmMyWork!lstCounseling.Visible And Not Forms!frmMyWork!lstWork.Visible, "Counseling tab did not switch views"
    found = False
    For i = 0 To Forms!frmMyWork!lstCounseling.ListCount - 1
        If CLng(Forms!frmMyWork!lstCounseling.ItemData(i)) = blank Then found = True
    Next
    Check Not found, "Member without a grade/rate listed in My Work"
    Check InStr(Forms!frmMyWork!sort_lstCounseling_1.Caption, "Member") = 1 And InStr(Forms!frmMyWork!sort_lstCounseling_3.Caption, "Status") = 1, "Counseling columns lack headings: " & Forms!frmMyWork!sort_lstCounseling_1.Caption
    Forms!frmMyWork!chkCounselingAll = True: UXMyWorkCounseling
    Forms!frmMyWork!lstCounseling = blank
    Check UXMyWorkCounselingRecord() And CurrentProject.AllForms("frmUXCounselingEditor").IsLoaded, "Record from My Work failed"
    UXDirectoryClose "frmUXCounselingEditor"
    Forms!frmMyWork!chkCounselingAll = True: UXMyWorkCounseling
    Check Forms!frmMyWork!lstCounseling.ListCount >= 3, "Show all did not list every member"
    Check UXMyWorkMode("Assignments") And Not Forms!frmMyWork!lstCounseling.Visible, "Leaving the tab kept counseling controls"
    DoCmd.Close acForm, "frmMyWork"

    stage = "disposition"
    Check Not CounselingDispositionDue(e5), "Disposition due before detachment"
    second = CommitDraft(0, "EVAL", "Detachment", "Synthetic member", #10/1/2026#, #10/20/2026#, "", "Test", "")
    Set s = CreateObject("Scripting.Dictionary"): s.Add "OccasionDetachment", "Yes"
    Call CommitSavedFields(second, CurrentRevision(second), s, "Detachment")
    LinkReportDirectoryMember second, e5, CurrentRevision(second), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & e5)), "Synthetic identity"
    Check CounselingDispositionDue(e5), "Disposition not due after detachment"
    Check InStr(CounselingWorkRows(True, #10/25/2026#), "Worksheet disposition due") > 0, "My Work did not list the disposition"
    Call SaveCounseling(id, 2, e5, #9/21/2026#, "Completed", "EXAMPLE, D", "", "Given to member", Null)
    Check Not CounselingDispositionDue(e5) And DLookup("DispositionOn", "MemberCounseling", "CounselingID=" & id) = Date, "Disposition not recorded with today's date"

    stage = "creation default"
    other = Member("Synthetic New Report", "IT2")
    Call SaveCounseling(0, 0, other, #4/10/2026#, "Completed", "EXAMPLE, F", "", "", Null)
    Set s = CreateObject("Scripting.Dictionary"): s("PeriodStart") = #9/16/2025#: s("PeriodEnd") = #9/15/2026#: s("CounselingDate") = Null: s("CounselorName") = Null
    CounselingCreationDefaults s, other
    Check s("CounselingDate") = "2026-04-10" And s("CounselorName") = "EXAMPLE, F", "New report did not start with the recorded session: " & s("CounselingDate")
    s("CounselingDate") = "2026-09-30": s("CounselorName") = "TYPED, BY USER"
    CounselingCreationDefaults s, other
    Check s("CounselingDate") = "2026-09-30" And s("CounselorName") = "TYPED, BY USER", "Creation default replaced an entered value"
    s("CounselingDate") = Null: s("CounselorName") = Null: s("ReportConcurrent") = "Yes"
    CounselingCreationDefaults s, other
    Check IsNull(s("CounselingDate")), "Concurrent report got a counseling default"

    stage = "merge"
    source = Member("Synthetic Duplicate", "IT2"): target = Member("Synthetic Survivor", "IT2")
    other = SaveCounseling(0, 0, source, #9/15/2026#, "Completed", "EXAMPLE, M", "", "", Null)
    stamp = PreviewDirectoryMerge(source, target)
    merge = ApplyDirectoryMerge(source, target, stamp, "Synthetic duplicate review")
    Check DLookup("MemberID", "MemberCounseling", "CounselingID=" & other) = target And DCount("*", "DirectoryMergeCounseling", "MergeID=" & merge) = 1, "Merge did not move counseling"
    Check UndoDirectoryMerge(merge), "Undo merge failed"
    Check DLookup("MemberID", "MemberCounseling", "CounselingID=" & other) = source, "Undo did not restore counseling"

    TestCounseling = "PASS: " & checks & " counseling schedule, status, record, carry-over, screen, disposition and merge assertions": Exit Function
Failed:
    TestCounseling = "FAIL at " & stage & " after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Public Function VerifyCounselingRecovery() As String
    Dim e5 As Long
    e5 = Nz(DLookup("MemberID", "CommandMember", "FullName='Synthetic Counseled'"), 0)
    If e5 = 0 Then VerifyCounselingRecovery = "FAIL: member missing": Exit Function
    If DLookup("WorksheetDisposition", "MemberCounseling", "MemberID=" & e5 & " AND IsArchived=False AND Outcome='Completed'") <> "Given to member" Then VerifyCounselingRecovery = "FAIL: disposition lost": Exit Function
    If DCount("*", "MemberCounseling", "MemberID=" & e5) <> 3 Then VerifyCounselingRecovery = "FAIL: record count " & DCount("*", "MemberCounseling", "MemberID=" & e5): Exit Function
    If DCount("*", "MemberCounseling", "MemberID=" & e5 & " AND NotPerformedReason='FAMILY EMERGENCY' AND IsArchived=False AND VersionCounter=3") <> 1 Then VerifyCounselingRecovery = "FAIL: restored counseling state/version lost": Exit Function
    VerifyCounselingRecovery = "PASS: counseling records, archive state and disposition persist in a new session"
End Function
