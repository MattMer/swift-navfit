Attribute VB_Name = "DirectoryHistoryHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6980, , message
    n = n + 1
End Sub
Private Function RejectMember(ByVal id As Long, ByVal version As Long, ByVal ssn As String) As Boolean
    On Error GoTo Rejected
    SaveDirectoryMember id, version, "TEST MEMBER", "TEST MEMBER", "IT2", ssn: Exit Function
Rejected:
    RejectMember = True
End Function
Public Function TestDirectoryHistory() As String
    Dim twinA As Long, twinB As Long, twinReport As Long, twinCopy As Long
    Dim m As Long, other As Long, id As Long, unknown As Long, version As Long, rev As Long, prior As Long, tied As Long, marks As Object, key As Variant, block As Long, previous As Object, count As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC COMMAND DIRECTORY", "Directory tester"
    m = SaveDirectoryMember(0, 0, "Casey Morgan", "MORGAN, CASEY", "IT2", "000000005")
    Check DLookup("MemberSSN", "CommandMember", "MemberID=" & m) = "000000005", "SSN leading zeros lost"
    Check RejectMember(0, 0, "000000005"), "Duplicate SSN allowed"
    Check RejectMember(m, 0, "000000005"), "Stale profile save allowed"
    other = SaveDirectoryMember(0, 0, "Casey Morgan", "MORGAN, CASEY", "IT2", "000000006")
    Check other <> m, "Name-only identity match merged people"
    id = CommitDraft(0, "EVAL", "Known member", "Casey Morgan", #1/1/2026#, #9/30/2026#, "", "Test", "", administration:=Array("MORGAN, CASEY", "IT2", "", "", ""), memberDetails:=Array("000000005", "", "Yes", "No", "No", "No", "2026-01-01"))
    Check DLookup("MemberID", "MemberReport", "ReportID=" & id) = m And DLookup("MatchState", "MemberReport", "ReportID=" & id) = "Confirmed", "Known SSN/report-name identity not linked"
    rev = CurrentRevision(id): count = DCount("*", "CommandMember")
    Call CommitDraft(id, "EVAL", "Known member", "Casey Morgan", #1/1/2026#, #9/30/2026#, "", "No-op", "")
    Check CurrentRevision(id) = rev And DCount("*", "CommandMember") = count, "No-op save created duplicate member/revision"
    unknown = CommitDraft(0, "EVAL", "Provisional", "New Person", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Check DLookup("MatchState", "MemberReport", "ReportID=" & unknown) = "Provisional", "Missing SSN was guessed by name"
    Check DCount("*", "CommandMember") = count + 1, "Unknown person not added to directory"
    Set previous = CheckedPrevious(unknown)
    Set previous = CheckedPrevious(id)
    Set marks = CreateObject("Scripting.Dictionary")
    For block = 33 To 39: marks.Add "Trait" & block, "4.0": Next
    prior = SavePriorEvaluation(0, 0, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Member-provided prior evaluation")
    Check PriorEvaluationITA(prior) = "4.00", "Prior ITA incorrect"
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "4.00" And previous("Promotion") = "Promotable", "Previous history lookup incorrect"
    Check previous("Period") = "2025-01-01 to 2025-12-31", "Prior period absent"
    Check CurrentRevision(id) = rev, "Prior entry changed current report revision"
    marks("Trait33") = "NOB"
    prior = SavePriorEvaluation(prior, 1, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Member-provided prior evaluation")
    Set previous = CheckedPrevious(id)
    Check PriorEvaluationITA(prior) = "4.00", "NOB included in prior numeric denominator"
    marks("Trait34") = Null
    prior = SavePriorEvaluation(prior, 2, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Incomplete copy")
    Set previous = CheckedPrevious(id)
    Check PriorEvaluationITA(prior) = "Incomplete", "Incomplete prior traits presented as final ITA"
    prior = SavePriorEvaluation(prior, 3, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "NOB", "NOB", marks, "NOB source")
    Check PriorEvaluationITA(prior) = "NOB", "Whole-report NOB not distinguished"
    Check DCount("*", "PriorEvaluationTrait", "PriorID=" & prior) = 7, "Trait edit lost slots or duplicated rows"
    LinkReportDirectoryMember unknown, m, CurrentRevision(unknown), 1, "Confirmed member identity from source"
    Check DLookup("MatchState", "MemberReport", "ReportID=" & unknown) = "Confirmed", "Manual identity confirmation failed"
    Set previous = CheckedPrevious(unknown)
    Check previous("ITA") = "NOB", "Confirmed identity did not expose prior history"
    For block = 33 To 39: marks("Trait" & block) = "5.0": Next
    tied = SavePriorEvaluation(0, 0, m, "FITREP", "ENS", #1/1/2025#, #12/31/2025#, "Must Promote", "Observed", marks, "Conflicting latest copy")
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "Ambiguous" And previous("Promotion") = "Ambiguous", "Equal-date history selected an arbitrary source"
    CurrentDb.Execute "UPDATE PriorEvaluation SET IsArchived=True WHERE PriorID=" & tied, dbFailOnError
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "NOB", "Archived external history remained eligible"
    tied = SavePriorEvaluation(0, 0, m, "EVAL", "IT2", #1/1/2026#, #1/1/2026#, "Must Promote", "Observed", marks, "Same-day source")
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "NOB", "History ending on current start was selected"
    ' Timestamp boundaries retain the original calendar-date selection semantics.
    CurrentDb.Execute "UPDATE Report SET PeriodStart=#2026-01-01 12:00:00# WHERE ReportID=" & id, dbFailOnError
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "NOB", "Same-day clock value selected prior history"
    CurrentDb.Execute "UPDATE Report SET PeriodStart=Null WHERE ReportID=" & id, dbFailOnError
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "Unavailable", "Missing start date selected history"
    CurrentDb.Execute "UPDATE Report SET PeriodStart=#2026-01-01# WHERE ReportID=" & id, dbFailOnError
    ' An archived Swift report is prior history only if it was ever marked signed;
    ' an abandoned, never-signed draft is not.
    CurrentDb.Execute "UPDATE Report SET PeriodStart=#2025-01-01#,PeriodEnd=#2025-12-31 12:00:00#,IsArchived=True WHERE ReportID=" & unknown, dbFailOnError
    Set previous = CheckedPrevious(id)
    Check previous("ITA") = "NOB", "Archived never-signed draft was used as the prior evaluation"
    CurrentDb.Execute "UPDATE Report SET IsArchived=False WHERE ReportID=" & unknown, dbFailOnError
    MarkReportSigned unknown, CurrentRevision(unknown), #1/15/2026#
    CurrentDb.Execute "UPDATE Report SET IsArchived=True WHERE ReportID=" & unknown, dbFailOnError
    Set previous = CheckedPrevious(id)
    Check Left$(previous("Source"), 18) = "Saved Swift report", "Archived signed report was not eligible as prior history"
    ' Duplicate keeps the original's member link; a same-name directory must not
    ' gain a third profile through fresh name matching.
    twinA = SaveDirectoryMember(0, 0, "Pat Lee", "LEE, PAT", "IT2", "")
    twinB = SaveDirectoryMember(0, 0, "Pat Lee", "LEE, PAT", "IT2", "")
    twinReport = CommitDraft(0, "EVAL", "Twin report", "Pat Lee", #1/1/2026#, #9/30/2026#, "", "Test", "", administration:=Array("LEE, PAT", "IT2", "", "", ""))
    LinkReportDirectoryMember twinReport, twinB, CurrentRevision(twinReport), DLookup("VersionCounter", "CommandMember", "MemberID=" & twinB), "Selected the right Pat Lee"
    count = DCount("*", "CommandMember")
    twinCopy = DuplicateDraft(twinReport)
    Check DLookup("MemberID", "MemberReport", "ReportID=" & twinCopy) = twinB And DLookup("MatchState", "MemberReport", "ReportID=" & twinCopy) = "Confirmed", "Duplicate lost the member link"
    Check DCount("*", "CommandMember") = count, "Duplicate created another member profile"
    TestDirectoryHistory = "PASS: " & n & " directory/history identity and prior-trait assertions": Exit Function
Failed:
    TestDirectoryHistory = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function

Private Function CheckedPrevious(ByVal id As Long) As Object
    Dim expected As Object, actual As Object, batch As Object, key As Variant
    Set expected = PreviousMemberEvaluation(id)
    Set batch = PreviousMemberEvaluationBatch("SELECT ReportID FROM Report")
    Set actual = batch(CStr(id))
    For Each key In expected.Keys
        Check CStr(expected(key)) = CStr(actual(key)), "Batch history differs: " & key
    Next
    Set CheckedPrevious = expected
End Function
