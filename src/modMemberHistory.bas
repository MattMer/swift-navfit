Attribute VB_Name = "modMemberHistory"
Option Compare Binary
Option Explicit
Public Function SavePriorEvaluation(ByVal priorID As Long, ByVal expectedVersion As Long, ByVal memberID As Long, ByVal kind As String, ByVal rank As String, ByVal fromDate As Date, ByVal toDate As Date, ByVal promotion As String, ByVal observation As String, ByVal traits As Object, ByVal sourceNote As String) As Long
    Dim r As DAO.Recordset, mark As Variant, block As Long, w As DAO.Workspace, author As Long, failure As Long, message As String, q As DAO.QueryDef, op As Long
    author = MajorAuthor()
    If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Choose a supported prior report type."
    If toDate < fromDate Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Prior report end precedes its start."
    If Len(rank) > 5 Or Len(sourceNote) > 255 Or Trim$(sourceNote) = "" Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Enter source/context within 255 characters and rank/rate within 5."
    If observation <> "Observed" And observation <> "NOB" Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Choose observed or whole-report NOB."
    If promotion <> "" Then mark = NAVFITEncode(promotion, "promotion")
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Choose an active command member."
    For block = 33 To 39
        If Not traits.Exists("Trait" & block) Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Include all seven trait slots; unentered values may remain blank."
        If Not IsNull(traits("Trait" & block)) And Nz(traits("Trait" & block), "") <> "" Then mark = NAVFITEncode(traits("Trait" & block), "trait")
    Next
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & priorID, dbOpenDynaset)
    If priorID = 0 Then
        r.AddNew: r!PriorGuid = NewIdentity(): r!CreatedBy = author: r!CreatedAt = Now(): r!VersionCounter = 1: r!IsArchived = False
    Else
        If r.EOF Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Prior record no longer exists."
        'Lock the row before checking its version, so a save another session is finishing is seen.
        r.Edit
        If r!MemberID <> memberID Or r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Prior record changed. Reopen it before editing."
        If r!IsArchived Then Err.Raise vbObjectError + 6970, "modMemberHistory.SavePriorEvaluation", "Archived prior history is read only."
        r!VersionCounter = r!VersionCounter + 1
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
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modMemberHistory", "SavePriorEvaluation", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not q Is Nothing Then SwiftClose q, swiftPrimary, "q"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function PriorEvaluationITARows(ByVal memberID As Long) As Object
    Dim result As Object, facts As Object, r As DAO.Recordset, id As Variant, row As Variant, mark As String
    Set result = CreateObject("Scripting.Dictionary"): Set facts = CreateObject("Scripting.Dictionary")
    Set r = CurrentDb.OpenRecordset("SELECT PriorID,ObservationStatus FROM PriorEvaluation WHERE MemberID=" & memberID & " AND IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        facts.Add CStr(r!PriorID), Array(Nz(r!ObservationStatus, ""), 0&, 0&, 0#): r.MoveNext
    Loop
    r.Close
    Set r = CurrentDb.OpenRecordset("SELECT t.PriorID,t.MarkValue FROM PriorEvaluationTrait AS t INNER JOIN PriorEvaluation AS p ON t.PriorID=p.PriorID WHERE p.MemberID=" & memberID & " AND p.IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        id = CStr(r!PriorID): row = facts(id): mark = Nz(r!MarkValue, "")
        If mark <> "" Then row(1) = row(1) + 1
        If mark <> "NOB" And mark <> "" Then row(2) = row(2) + 1: row(3) = row(3) + Val(mark)
        facts(id) = row: r.MoveNext
    Loop
    r.Close
    For Each id In facts.Keys
        row = facts(id)
        If row(0) = "NOB" Then
            result.Add CStr(id), "NOB"
        ElseIf row(1) < 7 Then
            result.Add CStr(id), "Incomplete"
        ElseIf row(2) = 0 Then
            result.Add CStr(id), "NOB"
        Else
            result.Add CStr(id), Format$(row(3) / row(2), "0.00")
        End If
    Next
    Set PriorEvaluationITARows = result
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
' Continuity across a member's whole record: prior evaluations and confirmed, active
' Swift reports that may be Regular, in date order. Prior evaluations do not record
' Regular or Concurrent, so a gap is only possible; an overlap is flagged only between
' two Swift Regular reports. Identical periods count once (the same report recorded twice,
' or alternative drafts). Rows are Array(kind, from, to, days), kind "Gap" or "Overlap".
Public Function MemberContinuityRows(ByVal memberID As Long) As Collection
    Dim rows As New Collection, r As DAO.Recordset, n As Long, i As Long, j As Long, found As Boolean
    Dim starts() As Date, ends() As Date, regular() As Boolean, s As Date, e As Date, b As Boolean, latest As Date
    Set MemberContinuityRows = rows
    Set r = CurrentDb.OpenRecordset("SELECT PeriodStart,PeriodEnd,False AS IsRegular FROM PriorEvaluation WHERE MemberID=" & memberID & " AND IsArchived=False AND PeriodStart Is Not Null AND PeriodEnd Is Not Null " & _
        "UNION ALL SELECT r.PeriodStart,r.PeriodEnd,IIf(r.ReportRegular='Yes',True,False) FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND m.MatchState='Confirmed' AND r.IsArchived=False AND r.PeriodStart Is Not Null AND r.PeriodEnd Is Not Null " & _
        "AND (r.ReportRegular='Yes' OR (Nz(r.ReportConcurrent,'')<>'Yes' AND Nz(r.ReportOperationalCommander,'')<>'Yes'))", dbOpenSnapshot)
    Do While Not r.EOF
        s = r!PeriodStart: e = r!PeriodEnd: b = CBool(r!IsRegular): found = False
        For i = 0 To n - 1
            If starts(i) = s And ends(i) = e Then regular(i) = regular(i) Or b: found = True: Exit For
        Next
        If Not found Then
            ReDim Preserve starts(0 To n): ReDim Preserve ends(0 To n): ReDim Preserve regular(0 To n)
            'Insert in order of start, then end.
            For i = n To 1 Step -1
                If starts(i - 1) < s Or (starts(i - 1) = s And ends(i - 1) <= e) Then Exit For
                starts(i) = starts(i - 1): ends(i) = ends(i - 1): regular(i) = regular(i - 1)
            Next
            starts(i) = s: ends(i) = e: regular(i) = b: n = n + 1
        End If
        r.MoveNext
    Loop
    r.Close
    If n < 2 Then Exit Function
    latest = ends(0)
    For i = 1 To n - 1
        If starts(i) > latest + 1 Then rows.Add Array("Gap", latest + 1, starts(i) - 1, CLng(starts(i) - latest - 1))
        If ends(i) > latest Then latest = ends(i)
    Next
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If regular(i) And regular(j) And starts(j) <= ends(i) Then rows.Add Array("Overlap", starts(j), IIf(ends(j) < ends(i), ends(j), ends(i)), 0)
        Next
    Next
End Function
Public Function MemberContinuitySummary(ByVal rows As Collection) As String
    Dim item As Variant, gaps As Long, overlaps As Long, text As String
    For Each item In rows
        If item(0) = "Gap" Then
            gaps = gaps + 1
        Else
            overlaps = overlaps + 1
        End If
    Next
    If gaps + overlaps = 0 Then MemberContinuitySummary = "Continuity: no gaps found.": Exit Function
    If gaps > 0 Then text = gaps & " possible gap" & IIf(gaps > 1, "s", "")
    If overlaps > 0 Then text = text & IIf(Len(text) > 0, ", ", "") & overlaps & " overlap" & IIf(overlaps > 1, "s", "") & " between Regular reports"
    MemberContinuitySummary = "Continuity: " & text & "." & IIf(gaps > 0, " Prior reports do not record Regular or Concurrent; verify on BOL Continuity.", "")
End Function
Public Function PreviousMemberEvaluation(ByVal reportID As Long) As Object
    Dim result As Object, member As Variant, start As Variant, r As DAO.Recordset, sql As String, source As String, recordID As Long, ending As Date, priorITA As String
    Set result = CreateObject("Scripting.Dictionary"): result.Add "ITA", "Unavailable": result.Add "Promotion", "Unavailable": result.Add "Period", "": result.Add "Source", "No earlier evaluation recorded"
    member = MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & reportID & " AND MatchState='Confirmed'")
    start = MajorScalar("SELECT PeriodStart FROM Report WHERE ReportID=" & reportID)
    If IsNull(member) Or IsNull(start) Then result("Source") = "Confirm member identity and report start date to select prior history": Set PreviousMemberEvaluation = result: Exit Function
    'An archived Swift report counts only if it was ever marked signed: an
    'abandoned draft is not an earlier evaluation.
    sql = "SELECT 'External' AS SourceKind,PriorID AS SourceID,PeriodStart,PeriodEnd,PromotionRecommendation FROM PriorEvaluation WHERE MemberID=" & CLng(member) & " AND IsArchived=False AND PeriodEnd<" & SQLDate(start)
    sql = sql & " UNION ALL SELECT 'Swift',r.ReportID,r.PeriodStart,r.PeriodEnd,r.PromotionRecommendation FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE (r.IsArchived=False OR r.ReportID In (SELECT ReportID FROM ReportSignoff)) AND m.MemberID=" & CLng(member) & " AND m.MatchState='Confirmed' AND r.ReportID<>" & reportID & " AND r.PeriodEnd<" & SQLDate(start) & " ORDER BY PeriodEnd DESC"
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

' Batch identity and candidate-history reads for one grid refresh. Selected ITA still
' uses its source-specific calculation; ties use the existing ambiguity behavior.
Public Function PreviousMemberEvaluationBatch(ByVal reportSQL As String) As Object
    Dim db As DAO.Database, r As DAO.Recordset, contexts As Object, candidates As Object, results As Object, computed As Object
    Dim context As Variant, item As Variant, first As Variant, id As Variant, memberSQL As String, sql As String, token As String
    Dim list As Collection, result As Object, selected As Boolean, tied As Boolean, saved As DAO.Recordset
    Set db = CurrentDb: Set contexts = CreateObject("Scripting.Dictionary"): Set candidates = CreateObject("Scripting.Dictionary")
    Set results = CreateObject("Scripting.Dictionary"): Set computed = CreateObject("Scripting.Dictionary")
    Set r = db.OpenRecordset("SELECT r.ReportID,r.PeriodStart,m.MemberID,m.MatchState FROM Report AS r LEFT JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE r.ReportID In (" & reportSQL & ")", dbOpenSnapshot)
    Do While Not r.EOF
        If Nz(r!MatchState, "") = "Confirmed" Then contexts(CStr(r!ReportID)) = Array(r!MemberID.Value, r!PeriodStart.Value) Else contexts(CStr(r!ReportID)) = Array(Null, r!PeriodStart.Value)
        r.MoveNext
    Loop
    r.Close
    memberSQL = "SELECT MemberID FROM MemberReport WHERE MatchState='Confirmed' AND ReportID In (" & reportSQL & ")"
    'Swift candidates carry what UXTableITA reads: observation status and head revision.
    sql = "SELECT 'External' AS SourceKind,PriorID AS SourceID,MemberID,PeriodStart,PeriodEnd,PromotionRecommendation,SourceNote AS SourceDetail,Null AS Observation,0 AS HeadRevision FROM PriorEvaluation WHERE IsArchived=False AND MemberID In (" & memberSQL & ")"
    sql = sql & " UNION ALL SELECT 'Swift',r.ReportID,m.MemberID,r.PeriodStart,r.PeriodEnd,r.PromotionRecommendation,r.WorkflowState,r.ObservationStatus,Nz(h.RevisionID,0) FROM (Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID) LEFT JOIN ReportHead AS h ON r.ReportID=h.ReportID WHERE (r.IsArchived=False OR r.ReportID In (SELECT ReportID FROM ReportSignoff)) AND m.MatchState='Confirmed' AND m.MemberID In (" & memberSQL & ") ORDER BY PeriodEnd DESC"
    Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not r.EOF
        token = CStr(r!MemberID)
        If Not candidates.Exists(token) Then Set list = New Collection: candidates.Add token, list
        Set list = candidates(token)
        list.Add Array(r!SourceKind.Value, r!SourceID.Value, r!PeriodStart.Value, r!PeriodEnd.Value, r!PromotionRecommendation.Value, r!SourceDetail.Value, r!Observation.Value, r!HeadRevision.Value)
        r.MoveNext
    Loop
    r.Close
    For Each id In contexts.Keys
        context = contexts(id): selected = False: tied = False
        Set result = CreateObject("Scripting.Dictionary"): result.Add "ITA", "Unavailable": result.Add "Promotion", "Unavailable": result.Add "Period", "": result.Add "Source", "No earlier evaluation recorded"
        If IsNull(context(0)) Or IsNull(context(1)) Then
            result("Source") = "Confirm member identity and report start date to select prior history"
        ElseIf candidates.Exists(CStr(context(0))) Then
            Set list = candidates(CStr(context(0)))
            For Each item In list
                If Not IsNull(item(3)) Then
                    If item(3) < DateValue(context(1)) And Not (item(0) = "Swift" And CLng(item(1)) = CLng(id)) Then
                        If Not selected Then
                            first = item: selected = True
                        Else
                            tied = (item(3) = first(3)): Exit For
                        End If
                    End If
                End If
            Next
            If tied Then
                Set result = PreviousMemberEvaluation(CLng(id))
            ElseIf selected Then
                token = first(0) & ":" & first(1)
                If Not computed.Exists(token) Then
                    If first(0) = "External" Then
                        computed(token) = PriorEvaluationITA(CLng(first(1)))
                    ElseIf CLng(first(7)) = 0 Then
                        computed(token) = UXTableITA(CLng(first(1)))
                    Else
                        If saved Is Nothing Then Set saved = SwiftDataDb(db).OpenRecordset("RevisionValue", dbOpenTable): saved.Index = "UX_RevisionValue"
                        computed(token) = SavedTraitITA(saved, first(6), CLng(first(7)))
                    End If
                End If
                result("ITA") = computed(token): result("Promotion") = Nz(first(4), "Unavailable")
                result("Period") = Format$(first(2), "yyyy-mm-dd") & " to " & Format$(first(3), "yyyy-mm-dd")
                If first(0) = "External" Then result("Source") = "Prior record: " & first(5) Else result("Source") = "Saved Swift report (local state " & first(5) & ")"
            End If
        End If
        results.Add CStr(id), result
    Next
    If Not saved Is Nothing Then saved.Close
    Set PreviousMemberEvaluationBatch = results
End Function
'UXTableITA for a known head revision, reading its saved traits by index seek.
Private Function SavedTraitITA(ByVal saved As DAO.Recordset, ByVal observation As Variant, ByVal revisionID As Long) As String
    Dim key As Variant, value As Variant, total As Double, count As Long
    Dim sparse As Boolean
    If Nz(observation, "") = "NOB" Then SavedTraitITA = "NOB": Exit Function
    sparse = RevisionIsSparse(revisionID)
    For Each key In TraitKeys()
        saved.Seek "=", revisionID, CStr(key)
        If saved.NoMatch Then
            'A sparse snapshot leaves an empty trait out; a full copy without it is incomplete.
            If Not sparse Then SavedTraitITA = "Unavailable": Exit Function
        ElseIf Not saved!IsNullValue Then
            value = saved!ValueText.Value
            If Not IsNull(value) Then
                If value <> "" And value <> "NOB" Then total = total + Val(value): count = count + 1
            End If
        End If
    Next
    If count = 0 Then SavedTraitITA = "Unavailable" Else SavedTraitITA = Format$(total / count, "0.00")
End Function
