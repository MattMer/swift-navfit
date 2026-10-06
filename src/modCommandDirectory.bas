Attribute VB_Name = "modCommandDirectory"
Option Compare Binary
Option Explicit
Private creationMember As Long
Private Function DirectorySSN(ByVal value As Variant) As String
    Dim s As String, i As Long
    s = Trim$(Nz(value, ""))
    If Len(s) <> 9 Then Exit Function
    For i = 1 To 9: If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    DirectorySSN = s
End Function
Private Function DirectoryName(ByVal value As Variant) As String
    Dim s As String
    s = UCase$(Trim$(Nz(value, "")))
    s = Replace(Replace(Replace(s, ",", ""), ".", ""), " ", "")
    DirectoryName = s
End Function
Public Function SaveDirectoryMember(ByVal memberID As Long, ByVal expectedVersion As Long, ByVal displayName As String, ByVal reportName As String, ByVal rank As String, ByVal ssn As String, Optional ByVal sourceKind As String = "Manual") As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, author As Long, failure As Long, message As String, op As Long, duplicate As Variant
    author = MajorAuthor(): displayName = Trim$(displayName): ssn = Trim$(ssn)
    If displayName = "" Or Len(displayName) > 120 Then Err.Raise vbObjectError + 6960, , "Enter a member name within 120 characters."
    If Len(reportName) > 27 Or Len(rank) > 5 Then Err.Raise vbObjectError + 6960, , "Report name permits 27 characters; rank/rate permits 5. No value will be truncated."
    If ssn <> "" And DirectorySSN(ssn) = "" Then Err.Raise vbObjectError + 6960, , "Enter all nine SSN digits or leave SSN unentered."
    If ssn <> "" Then
        duplicate = MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn) & " AND MemberID<>" & memberID)
        If Not IsNull(duplicate) Then Err.Raise vbObjectError + 6960, , "That SSN already belongs to a directory member. Choose the existing member; names alone are not an identity match."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenDynaset)
    If memberID = 0 Then
        r.AddNew: r!MemberGuid = NewIdentity(): r!CreatedBy = author: r!CreatedAt = Now(): r!SourceKind = sourceKind: r!IsArchived = False: r!VersionCounter = 1
    Else
        If r.EOF Then Err.Raise vbObjectError + 6960, , "Member no longer exists."
        If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6960, , "Member profile changed. Reopen it before saving."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!FullName = displayName: r!ReportName = reportName: r!RankRate = rank
    If ssn = "" Then r!MemberSSN = Null Else r!MemberSSN = ssn
    r!ModifiedBy = author: r!ModifiedAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveDirectoryMember = r!MemberID: r.Close
    op = BeginOperation("Save command member", displayName & "; rank/rate " & rank & "; SSN " & IIf(ssn = "", "unentered", "ending " & Right$(ssn, 4)))
    FinishOperation op, "Completed", "Directory profile saved; report fields and prior evaluation records unchanged."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "SaveDirectoryMember", message
End Function
Public Function SyncReportDirectory(ByVal reportID As Long) As Long
    Dim report As DAO.Recordset, link As DAO.Recordset, member As DAO.Recordset, id As Long, ssn As String, name As String, official As String, state As String, source As String, existing As Variant
    Set report = CurrentDb.OpenRecordset("SELECT MemberName,OfficialName,MemberSSN,RankRate FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If report.EOF Then report.Close: Exit Function
    official = Nz(report!OfficialName, ""): name = UXIdentity(official, report!MemberName)
    If name = "Unnamed member" Then report.Close: Exit Function
    ssn = DirectorySSN(report!MemberSSN)
    Set link = CurrentDb.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & reportID, dbOpenDynaset)
    If creationMember > 0 Then
        id = creationMember: state = "Confirmed": source = "Manual selection"
    ElseIf Not link.EOF Then
        id = link!MemberID: state = link!MatchState: source = link!MatchSource
        Set member = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
        If ssn <> "" And ssn <> Nz(member!MemberSSN, "") Then
            state = "Needs review"
            existing = MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn))
            If IsNull(existing) And Nz(member!MemberSSN, "") = "" And source = "Report creation" Then
                SaveDirectoryMember id, member!VersionCounter, member!FullName, Nz(member!ReportName, ""), Nz(member!RankRate, ""), ssn, "Report"
                state = "Confirmed"
            End If
        End If
        member.Close
    Else
        If ssn <> "" Then existing = MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn)) Else existing = Null
        If IsNull(existing) Then
            id = SaveDirectoryMember(0, 0, name, official, Nz(report!RankRate, ""), ssn, "Report")
            state = IIf(ssn = "", "Provisional", "Confirmed"): source = "Report creation"
        Else
            id = CLng(existing): source = "SSN match"
            Set member = CurrentDb.OpenRecordset("SELECT FullName,ReportName FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
            state = "Needs review"
            If DirectoryName(name) = DirectoryName(member!FullName) Or (official <> "" And DirectoryName(official) = DirectoryName(member!ReportName)) Then state = "Confirmed"
            member.Close
        End If
    End If
    If link.EOF Then link.AddNew: link!ReportID = reportID Else link.Edit
    link!MemberID = id: link!MatchState = state: link!MatchSource = source: link!UpdatedBy = MajorAuthor(): link!UpdatedAt = Now(): link.Update
    link.Close: report.Close: SyncReportDirectory = id
End Function
Public Function SyncDirectoryFromReports() As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM Report ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        If SyncReportDirectory(r!ReportID) > 0 Then SyncDirectoryFromReports = SyncDirectoryFromReports + 1
        r.MoveNext
    Loop
    r.Close
End Function
Public Sub LinkReportDirectoryMember(ByVal reportID As Long, ByVal memberID As Long, ByVal expectedRevision As Long, ByVal expectedMemberVersion As Long, ByVal reason As String)
    Dim w As DAO.Workspace, r As DAO.Recordset, op As Long, failure As Long, message As String
    If Trim$(reason) = "" Then Err.Raise vbObjectError + 6961, , "Record why this is the correct member."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If CurrentRevision(reportID) <> expectedRevision Or Nz(MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & memberID), 0) <> expectedMemberVersion Then Err.Raise vbObjectError + 6961, , "Report or member changed. Preview the identity link again."
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 6961, , "Choose an active directory member."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & reportID, dbOpenDynaset)
    op = BeginOperation("Confirm report member identity", "Report " & reportID & "; prior directory member " & IIf(r.EOF, "unlinked", Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & reportID), 0)) & "; new member " & memberID & "; reason " & reason)
    If r.EOF Then r.AddNew: r!ReportID = reportID Else r.Edit
    r!MemberID = memberID: r!MatchState = "Confirmed": r!MatchSource = "Manual confirmation": r!UpdatedBy = MajorAuthor(): r!UpdatedAt = Now(): r.Update: r.Close
    FinishOperation op, "Completed", "Identity link confirmed; saved report fields unchanged."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "LinkReportDirectoryMember", message
End Sub
Public Sub DirectoryCreationContext(ByVal memberID As Long)
    creationMember = memberID
End Sub
