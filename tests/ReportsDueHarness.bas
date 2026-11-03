Attribute VB_Name = "ReportsDueHarness"
Option Explicit
'F03: reports due from the periodic calendar and departures, cleared by a draft.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7890, , message
    checks = checks + 1
End Sub
'The due item key for a member and basis, or "" when it is not listed.
Private Function DueKey(ByVal items As Collection, ByVal name As String, ByVal basis As String) As String
    Dim item As Variant
    For Each item In items
        If item(1) = name And item(3) = basis Then DueKey = item(0) & "|" & item(3) & "|" & Format$(item(5), "yyyy-mm-dd"): Exit Function
    Next
End Function
Private Function Refused(ByVal action As String, ByVal key As String) As Boolean
    On Error GoTo Failed
    If action = "create" Then Call CreateDueDraft(key) Else Call DismissDue(key, "")
    Exit Function
Failed:
    Refused = True
End Function
Private Function DateRefused(ByVal member As Long, ByVal value As String) As Boolean
    On Error GoTo Failed
    Call SaveDirectoryMember(member, DLookup("VersionCounter", "CommandMember", "MemberID=" & member), DLookup("FullName", "CommandMember", "MemberID=" & member), Nz(DLookup("ReportName", "CommandMember", "MemberID=" & member), ""), Nz(DLookup("RankRate", "CommandMember", "MemberID=" & member), ""), "", "Manual", , , value)
    Exit Function
Failed:
    DateRefused = True
End Function
Private Sub SetDateReported(ByVal member As Long, ByVal value As String)
    Call SaveDirectoryMember(member, DLookup("VersionCounter", "CommandMember", "MemberID=" & member), DLookup("FullName", "CommandMember", "MemberID=" & member), Nz(DLookup("ReportName", "CommandMember", "MemberID=" & member), ""), Nz(DLookup("RankRate", "CommandMember", "MemberID=" & member), ""), "", "Manual", , , value)
End Sub
Public Function TestReportsDue() As String
    Dim items As Collection, missing As Long, asOf As Date, e5 As Long, e6 As Long, o3 As Long, e9 As Long, chief As Long, key As String, id As Long, other As Long, changes As Object, dismissal As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC REPORTS DUE", "Coordinator"
    e5 = SaveDirectoryMember(0, 0, "EVANS FIVE", "EVANS, FIVE", "IT2", "")
    e6 = SaveDirectoryMember(0, 0, "EDWARDS SIX", "EDWARDS, SIX", "IT1", "")
    o3 = SaveDirectoryMember(0, 0, "OWENS THREE", "OWENS, THREE", "LT", "")
    e9 = SaveDirectoryMember(0, 0, "MASTER CHIEF", "CHIEF, MASTER", "ITCM", "")
    chief = SaveDirectoryMember(0, 0, "CHIEF SEVEN", "SEVEN, CHIEF", "ITC", "")
    Call SaveDirectoryMember(0, 0, "NO RANK", "RANK, NO", "", "")
    asOf = #2/1/2026#

    ' Periodic dates by paygrade, within 90 days ahead or 90 days overdue.
    Set items = ReportsDueItems(asOf, missing)
    Check missing = 1, "Member without rank/rate not counted"
    Check DueKey(items, "EVANS FIVE", "Periodic") = e5 & "|Periodic|2026-03-15", "E-5 periodic date wrong"
    Check DueKey(items, "EDWARDS SIX", "Periodic") = e6 & "|Periodic|2025-11-15", "Overdue E-6 periodic missing"
    Check DueKey(items, "OWENS THREE", "Periodic") = o3 & "|Periodic|2026-01-31", "O-3 periodic date wrong"
    Check DueKey(items, "MASTER CHIEF", "Periodic") = e9 & "|Periodic|2026-04-15", "E-9 periodic before NAVADMIN 132/26 missing"
    Check DueKey(items, "CHIEF SEVEN", "Periodic") = "", "E-7 periodic outside the window listed"
    Check items(1)(5) <= items(items.Count)(5), "Due list not soonest first"
    Check DueKey(ReportsDueItems(#7/1/2026#), "MASTER CHIEF", "Periodic") = "", "E-9 still due after NAVADMIN 132/26"
    Check DueFormKind("E6") = "EVAL" And DueFormKind("E7") = "CHIEFEVAL" And DueFormKind("O3") = "FITREP" And DueFormKind("W2") = "FITREP", "Report form by paygrade wrong"

    ' Create draft clears the item and links the member.
    key = DueKey(items, "EVANS FIVE", "Periodic")
    id = CreateDueDraft(key)
    Check DLookup("FormKind", "Report", "ReportID=" & id) = "EVAL" And DLookup("PeriodEnd", "Report", "ReportID=" & id) = #3/15/2026#, "Draft form or period wrong"
    Check DLookup("ReportRegular", "Report", "ReportID=" & id) = "Yes" And DLookup("OccasionPeriodic", "Report", "ReportID=" & id) = "Yes", "Draft occasion wrong"
    Check DLookup("MemberID", "MemberReport", "ReportID=" & id) = e5, "Draft not linked to the member"
    Check DueKey(ReportsDueItems(asOf), "EVANS FIVE", "Periodic") = "", "Draft did not clear the due report"
    Check Refused("create", key), "Second draft created for a satisfied report"

    ' A Concurrent-only report does not satisfy the Regular periodic report.
    other = CommitDraft(0, "EVAL", "Concurrent", "EDWARDS SIX", #1/1/2025#, #11/15/2025#, "")
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "ReportConcurrent", "Yes"
    Call CommitSavedFields(other, CurrentRevision(other), changes, "Concurrent only")
    LinkReportDirectoryMember other, e6, CurrentRevision(other), DLookup("VersionCounter", "CommandMember", "MemberID=" & e6), "Synthetic link"
    Check DueKey(ReportsDueItems(asOf), "EDWARDS SIX", "Periodic") <> "", "Concurrent-only report cleared the periodic report"

    ' Dismiss needs a reason; undo brings the item back.
    key = DueKey(ReportsDueItems(asOf), "OWENS THREE", "Periodic")
    Check Refused("dismiss", key), "Dismissed without a reason"
    dismissal = DismissDue(key, "Transferred before the periodic date")
    Check DueKey(ReportsDueItems(asOf), "OWENS THREE", "Periodic") = "", "Dismissed report still listed"
    UndoDueDismissal dismissal
    Check DueKey(ReportsDueItems(asOf), "OWENS THREE", "Periodic") <> "", "Undone dismissal not listed again"

    ' A departure owes a detachment report, cleared by its draft.
    CurrentDb.Execute "INSERT INTO MemberDeparture (MemberID,Basis,ExpectedDate,ConfirmedDate,DepartureState,CreatedBy,CreatedAt,VersionCounter) VALUES (" & e5 & ",'PRD',#2026-04-20#,#2026-04-20#,'Confirmed'," & MajorAuthor() & ",Now(),1)", dbFailOnError
    key = DueKey(ReportsDueItems(asOf), "EVANS FIVE", "Detachment")
    Check key = e5 & "|Detachment|2026-04-20", "Detachment report not due"
    id = CreateDueDraft(key)
    Check DLookup("OccasionDetachment", "Report", "ReportID=" & id) = "Yes" And DLookup("PeriodStart", "Report", "ReportID=" & id) = #3/16/2026#, "Detachment draft wrong or did not start after the last report"
    Check DueKey(ReportsDueItems(asOf), "EVANS FIVE", "Detachment") = "", "Detachment draft did not clear the item"
    Check InStr(ReportsDueRows(), "OWENS, THREE") = 0 And InStr(ReportsDueRows(), "OWENS THREE") > 0, "Due list rows wrong"

    ' Date reported: entered on the profile, copied into new drafts, moved forward by a saved report.
    Check DateRefused(e6, "not a date") And DateRefused(e6, Format$(DateAdd("yyyy", 2, Date), "yyyy-mm-dd")), "Invalid date reported accepted"
    SetDateReported e6, "2024-06-10"
    Check DLookup("DateReported", "CommandMember", "MemberID=" & e6) = #6/10/2024#, "Profile date reported not saved"
    id = CreateDueDraft(DueKey(ReportsDueItems(asOf), "EDWARDS SIX", "Periodic"))
    Check DLookup("DateReported", "Report", "ReportID=" & id) = "2024-06-10", "Draft did not take the profile's date reported"
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "DateReported", "2025-01-05"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Reported again")
    Check DLookup("DateReported", "CommandMember", "MemberID=" & e6) = #1/5/2025#, "A newer date reported on a report did not update the profile"
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "DateReported", "2023-01-01"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Older date")
    Check DLookup("DateReported", "CommandMember", "MemberID=" & e6) = #1/5/2025#, "An older report date moved the profile back"
    Call SaveDirectoryMember(e6, DLookup("VersionCounter", "CommandMember", "MemberID=" & e6), "EDWARDS SIX", "EDWARDS, SIX", "IT1", "")
    Check DLookup("DateReported", "CommandMember", "MemberID=" & e6) = #1/5/2025#, "A save without the date cleared it"
    SetDateReported e6, ""
    Check IsNull(DLookup("DateReported", "CommandMember", "MemberID=" & e6)), "Blank did not clear date reported"
    TestReportsDue = "PASS: " & checks & " reports-due calendar, draft, dismissal, detachment and date-reported assertions": Exit Function
Failed:
    TestReportsDue = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
