Attribute VB_Name = "DirectoryDemoSeed"
Option Explicit
Public Function SeedDirectoryDemoHistory() As String
    Dim r As DAO.Recordset, marks As Object, block As Long, id As Long, added As Long
    Dim kind As String, startDate As Date, endDate As Date, revisions As Long
    On Error GoTo Failed
    If InStr(1, Nz(DLookup("CommandName", "Workspace"), ""), "DEMO - SYNTHETIC DATA", vbTextCompare) = 0 Then Err.Raise vbObjectError + 6999, , "Only the synthetic demo may be seeded."
    revisions = DCount("*", "ReportRevision")
    Set r = CurrentDb.OpenRecordset("SELECT MemberID,RankRate FROM CommandMember WHERE IsArchived=False ORDER BY MemberID", dbOpenSnapshot)
    Do While Not r.EOF
        id = Nz(MajorScalar("SELECT TOP 1 r.ReportID FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & r!MemberID & " AND m.MatchState='Confirmed' AND r.PeriodStart Is Not Null ORDER BY r.PeriodStart"), 0)
        If id > 0 And DCount("*", "PriorEvaluation", "MemberID=" & r!MemberID) = 0 Then
            kind = DLookup("FormKind", "Report", "ReportID=" & id)
            endDate = DateAdd("d", -1, CDate(DLookup("PeriodStart", "Report", "ReportID=" & id)))
            startDate = DateAdd("yyyy", -1, DateAdd("d", 1, endDate))
            Set marks = CreateObject("Scripting.Dictionary")
            For block = 33 To 39: marks.Add "Trait" & block, IIf((block + r!MemberID) Mod 3 = 0, "3.0", "4.0"): Next
            Call SavePriorEvaluation(0, 0, r!MemberID, kind, Nz(r!RankRate, ""), startDate, endDate, "Promotable", "Observed", marks, "DEMO synthetic prior evaluation - not an actual Navy record")
            added = added + 1
        End If
        r.MoveNext
    Loop
    r.Close
    If DCount("*", "ReportRevision") <> revisions Then Err.Raise vbObjectError + 6999, , "Seeding changed current report revisions."
    SeedDirectoryDemoHistory = "PASS: added " & added & " synthetic prior evaluations; current revisions unchanged": Exit Function
Failed:
    SeedDirectoryDemoHistory = "FAIL: " & Err.Number & " " & Err.Description
End Function
