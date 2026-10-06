Attribute VB_Name = "modMemberHistory"
Option Compare Binary
Option Explicit
Public Function SavePriorEvaluation(ByVal priorID As Long, ByVal expectedVersion As Long, ByVal memberID As Long, ByVal kind As String, ByVal rank As String, ByVal fromDate As Date, ByVal toDate As Date, ByVal promotion As String, ByVal observation As String, ByVal traits As Object, ByVal sourceNote As String) As Long
    Dim r As DAO.Recordset, mark As Variant, block As Long, w As DAO.Workspace, author As Long, failure As Long, message As String, q As DAO.QueryDef, op As Long
    author = MajorAuthor()
    If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 6970, , "Choose a supported prior report type."
    If toDate < fromDate Then Err.Raise vbObjectError + 6970, , "Prior report end precedes its start."
    If Len(rank) > 5 Or Len(sourceNote) > 255 Or Trim$(sourceNote) = "" Then Err.Raise vbObjectError + 6970, , "Enter source/context within 255 characters and rank/rate within 5."
    If observation <> "Observed" And observation <> "NOB" Then Err.Raise vbObjectError + 6970, , "Choose observed or whole-report NOB."
    If promotion <> "" Then mark = NAVFITEncode(promotion, "promotion")
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 6970, , "Choose an active command member."
    For block = 33 To 39
        If Not traits.Exists("Trait" & block) Then Err.Raise vbObjectError + 6970, , "Include all seven trait slots; unentered values may remain blank."
        If Not IsNull(traits("Trait" & block)) And Nz(traits("Trait" & block), "") <> "" Then mark = NAVFITEncode(traits("Trait" & block), "trait")
    Next
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & priorID, dbOpenDynaset)
    If priorID = 0 Then
        r.AddNew: r!PriorGuid = NewIdentity(): r!CreatedBy = author: r!CreatedAt = Now(): r!VersionCounter = 1: r!IsArchived = False
    Else
        If r.EOF Then Err.Raise vbObjectError + 6970, , "Prior record no longer exists."
        If r!MemberID <> memberID Or r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6970, , "Prior record changed. Reopen it before editing."
        If r!IsArchived Then Err.Raise vbObjectError + 6970, , "Archived prior history is read only."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!MemberID = memberID: r!FormKind = kind: r!RankRate = rank: r!PeriodStart = DateValue(fromDate): r!PeriodEnd = DateValue(toDate)
    r!PromotionRecommendation = promotion: r!ObservationStatus = observation: r!SourceNote = sourceNote: r!ModifiedBy = author: r!ModifiedAt = Now()
    r.Update: r.Bookmark = r.LastModified: SavePriorEvaluation = r!PriorID: r.Close
    CurrentDb.Execute "DELETE FROM PriorEvaluationTrait WHERE PriorID=" & SavePriorEvaluation, dbFailOnError
    Set q = CurrentDb.CreateQueryDef("", "PARAMETERS id Long, block Short, mark Text(3); INSERT INTO PriorEvaluationTrait (PriorID,BlockNumber,MarkValue) VALUES (id,block,mark)")
    For block = 33 To 39: q!id = SavePriorEvaluation: q!block = block: q!mark = traits("Trait" & block): q.Execute dbFailOnError: Next
    q.Close
    op = BeginOperation("Save prior evaluation", "Member " & memberID & "; prior record " & SavePriorEvaluation & "; period " & Format$(fromDate, "yyyy-mm-dd") & " to " & Format$(toDate, "yyyy-mm-dd") & "; source " & sourceNote)
    FinishOperation op, "Completed", "Prior report fields and traits saved atomically. Current Swift reports unchanged."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    If Not q Is Nothing Then q.Close
    On Error GoTo 0: Err.Raise failure, "SavePriorEvaluation", message
End Function
Public Function PriorEvaluationITA(ByVal priorID As Long) As String
    Dim r As DAO.Recordset, count As Long, nob As Long, total As Double, mark As String, filled As Long
    If Nz(MajorScalar("SELECT ObservationStatus FROM PriorEvaluation WHERE PriorID=" & priorID), "") = "NOB" Then PriorEvaluationITA = "NOB": Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT MarkValue FROM PriorEvaluationTrait WHERE PriorID=" & priorID, dbOpenSnapshot)
    Do While Not r.EOF
        mark = Nz(r!MarkValue, "")
        If mark <> "" Then filled = filled + 1
        If mark = "NOB" Then nob = nob + 1 Else If mark <> "" Then total = total + Val(mark): count = count + 1
        r.MoveNext
    Loop
    r.Close
    If filled < 7 Then PriorEvaluationITA = "Incomplete": Exit Function
    If count = 0 Then PriorEvaluationITA = "NOB" Else PriorEvaluationITA = Format$(total / count, "0.00")
End Function
Public Function PreviousMemberEvaluation(ByVal reportID As Long) As Object
    Dim result As Object, member As Variant, start As Variant, r As DAO.Recordset, sql As String, source As String, recordID As Long, ending As Date, priorITA As String
    Set result = CreateObject("Scripting.Dictionary"): result.Add "ITA", "Unavailable": result.Add "Promotion", "Unavailable": result.Add "Period", "": result.Add "Source", "No earlier evaluation recorded"
    member = MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & reportID & " AND MatchState='Confirmed'")
    start = MajorScalar("SELECT PeriodStart FROM Report WHERE ReportID=" & reportID)
    If IsNull(member) Or IsNull(start) Then result("Source") = "Confirm member identity and report start date to select prior history": Set PreviousMemberEvaluation = result: Exit Function
    sql = "SELECT 'External' AS SourceKind,PriorID AS SourceID,PeriodStart,PeriodEnd,PromotionRecommendation FROM PriorEvaluation WHERE MemberID=" & CLng(member) & " AND IsArchived=False AND PeriodEnd<" & SQLDate(start)
    sql = sql & " UNION ALL SELECT 'Swift',r.ReportID,r.PeriodStart,r.PeriodEnd,r.PromotionRecommendation FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & CLng(member) & " AND m.MatchState='Confirmed' AND r.ReportID<>" & reportID & " AND r.PeriodEnd<" & SQLDate(start) & " ORDER BY PeriodEnd DESC"
    Set r = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not r.EOF Then
        source = r!SourceKind: recordID = r!SourceID: ending = r!PeriodEnd
        result("Promotion") = Nz(r!PromotionRecommendation, "Unavailable")
        result("Period") = Format$(r!PeriodStart, "yyyy-mm-dd") & " to " & Format$(r!PeriodEnd, "yyyy-mm-dd")
        If source = "External" Then result("ITA") = PriorEvaluationITA(recordID): result("Source") = "Prior record: " & MajorScalar("SELECT SourceNote FROM PriorEvaluation WHERE PriorID=" & recordID) Else result("ITA") = UXTableITA(recordID): result("Source") = "Saved Swift report (local state " & MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & recordID) & ")"
        r.MoveNext
        If Not r.EOF Then
            If r!PeriodEnd = ending Then result("ITA") = "Ambiguous": result("Promotion") = "Ambiguous": result("Source") = "Multiple reports share the latest prior ending date. Review member history; no source was guessed."
        End If
    End If
    r.Close: Set PreviousMemberEvaluation = result
End Function
Private Function SQLDate(ByVal value As Date) As String
    SQLDate = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function
