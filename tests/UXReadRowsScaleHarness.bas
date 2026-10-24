Attribute VB_Name = "UXReadRowsScaleHarness"
Option Explicit
Private Function ReadResult(ByVal sql As String) As String
    Dim db As DAO.Database, r As DAO.Recordset, f As DAO.Field
    Set db = CurrentDb: Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not r.EOF
        For Each f In r.Fields
            ReadResult = ReadResult & Len(Nz(f.Value, "")) & ":" & Nz(f.Value, "") & "|"
        Next
        ReadResult = ReadResult & vbLf: r.MoveNext
    Loop
    r.Close
End Function
Private Function Sample(ByVal prepared As Boolean, ByVal work As Boolean, ByRef result As String) As Double
    Dim started As Double, sql As String
    started = Timer
    If work Then
        If prepared Then UXPrepareWorkRows MajorAuthor(), False
        sql = UXWorkSQL(MajorAuthor(), False, prepared)
    ElseIf prepared Then
        UXPrepareReportRows
        sql = "SELECT ReportID,UXIdentity(OfficialName,MemberName),UXReportRowCell(ReportID,1),UXReportRowCell(ReportID,2),UXReportRowCell(ReportID,3),UXReportRowCell(ReportID,4) FROM Report ORDER BY ReportID"
    Else
        sql = "SELECT ReportID,UXMemberLabel(ReportID),UXReportState(ReportID),UXTableITA(ReportID),UXTableGroups(ReportID),DLookup('RevisionNumber','ReportRevision','RevisionID=' & CurrentRevisionID) FROM Report ORDER BY ReportID"
    End If
    result = ReadResult(sql): Sample = Timer - started
    If Sample < 0 Then Sample = Sample + 86400
    Sample = Sample * 1000
End Function
Public Function TestReadRowsScale(Optional ByVal maximumRows As Long = 100) As String
    Dim target As Variant, targets As Variant, i As Long, id As Long, p As Long, g As Long, template As Long, ignored As Long, work As Variant, trial As Long, before As String, after As String, a As Double, b As Double, output As String, revisions As Long, snapshots As Long
    On Error GoTo Failed
    If maximumRows < 100 Or maximumRows > 1000 Then Err.Raise 5, , "Use 100-1000 synthetic reports"
    ConfigureWorkspace "SYNTHETIC READ ROW SCALE", "Reviewer"
    p = MajorAuthor(): g = CreateSummaryGroup("Synthetic scale group", "EVAL", "E5")
    template = CreateRouteTemplate("Scale route", CStr(p), "Read")
    output = "Reports,View,Trial,UnpreparedMs,PreparedMs" & vbCrLf
    targets = Array(25, 100)
    If maximumRows > 100 Then targets = Array(25, 100, maximumRows)
    For Each target In targets
        For i = DCount("*", "Report") + 1 To CLng(target)
            id = CommitDraft(0, "EVAL", "Synthetic " & i, "MEMBER " & Format$(i, "000"), #1/1/2026#, #9/30/2026#, "", traits:=Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"))
            AddGroupMember g, id, "E5"
            ignored = AssignReview(id, CurrentRevision(id), p, Null, "Review")
            If i Mod 5 = 0 Then ignored = StartRoute(id, template, g)
            If i Mod 3 = 0 Then ignored = AddReviewComment(id, CurrentRevision(id), "Synthetic comment")
        Next
        revisions = DCount("*", "ReportRevision"): snapshots = DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet")
        For Each work In Array(False, True)
            For trial = 1 To 5
                If trial Mod 2 = 1 Then
                    a = Sample(False, CBool(work), before): b = Sample(True, CBool(work), after)
                Else
                    b = Sample(True, CBool(work), after): a = Sample(False, CBool(work), before)
                End If
                If StrComp(before, after, vbBinaryCompare) <> 0 Then Err.Raise 5, , "Scale row contents/order differ"
                output = output & target & "," & IIf(work, "MyWork", "Reports") & "," & trial & "," & Format$(a, "0.0") & "," & Format$(b, "0.0") & vbCrLf
            Next
        Next
        If DCount("*", "ReportRevision") <> revisions Or DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet") <> snapshots Then Err.Raise 5, , "Display reads created history"
    Next
    TestReadRowsScale = "PASS:" & vbCrLf & output: Exit Function
Failed:
    TestReadRowsScale = "FAIL: " & Err.Number & " " & Err.Description
End Function
