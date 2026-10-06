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
    Set marks = CreateObject("Scripting.Dictionary")
    For block = 33 To 39: marks.Add "Trait" & block, "4.0": Next
    prior = SavePriorEvaluation(0, 0, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Member-provided prior evaluation")
    Check PriorEvaluationITA(prior) = "4.00", "Prior ITA incorrect"
    Set previous = PreviousMemberEvaluation(id)
    Check previous("ITA") = "4.00" And previous("Promotion") = "Promotable", "Previous history lookup incorrect"
    Check previous("Period") = "2025-01-01 to 2025-12-31", "Prior period absent"
    Check CurrentRevision(id) = rev, "Prior entry changed current report revision"
    marks("Trait33") = "NOB"
    prior = SavePriorEvaluation(prior, 1, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Member-provided prior evaluation")
    Check PriorEvaluationITA(prior) = "4.00", "NOB included in prior numeric denominator"
    marks("Trait34") = Null
    prior = SavePriorEvaluation(prior, 2, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Incomplete copy")
    Check PriorEvaluationITA(prior) = "Incomplete", "Incomplete prior traits presented as final ITA"
    prior = SavePriorEvaluation(prior, 3, m, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "NOB", "NOB", marks, "NOB source")
    Check PriorEvaluationITA(prior) = "NOB", "Whole-report NOB not distinguished"
    Check DCount("*", "PriorEvaluationTrait", "PriorID=" & prior) = 7, "Trait edit lost slots or duplicated rows"
    LinkReportDirectoryMember unknown, m, CurrentRevision(unknown), 1, "Confirmed member identity from source"
    Check DLookup("MatchState", "MemberReport", "ReportID=" & unknown) = "Confirmed", "Manual identity confirmation failed"
    Set previous = PreviousMemberEvaluation(unknown)
    Check previous("ITA") = "NOB", "Confirmed identity did not expose prior history"
    For block = 33 To 39: marks("Trait" & block) = "5.0": Next
    tied = SavePriorEvaluation(0, 0, m, "FITREP", "ENS", #1/1/2025#, #12/31/2025#, "Must Promote", "Observed", marks, "Conflicting latest copy")
    Set previous = PreviousMemberEvaluation(id)
    Check previous("ITA") = "Ambiguous" And previous("Promotion") = "Ambiguous", "Equal-date history selected an arbitrary source"
    CurrentDb.Execute "UPDATE PriorEvaluation SET IsArchived=True WHERE PriorID=" & tied, dbFailOnError
    Set previous = PreviousMemberEvaluation(id)
    Check previous("ITA") = "NOB", "Archived external history remained eligible"
    tied = SavePriorEvaluation(0, 0, m, "EVAL", "IT2", #1/1/2026#, #1/1/2026#, "Must Promote", "Observed", marks, "Same-day source")
    Set previous = PreviousMemberEvaluation(id)
    Check previous("ITA") = "NOB", "History ending on current start was selected"
    TestDirectoryHistory = "PASS: " & n & " directory/history identity and prior-trait assertions": Exit Function
Failed:
    TestDirectoryHistory = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
