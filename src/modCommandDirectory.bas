Attribute VB_Name = "modCommandDirectory"
Option Compare Binary
Option Explicit
Private creationMember As Long, creationState As String, creationSource As String
Private Function DirectorySSN(ByVal value As Variant) As String
    Dim s As String, i As Long
    s = Trim$(Nz(value, ""))
    If Len(s) <> 9 Then Exit Function
    For i = 1 To 9: If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    DirectorySSN = s
End Function
 ' SSN is authoritative. Name fallback is exact after punctuation/spacing
 ' normalization, unique among active members, and never crosses conflicting SSNs.
Public Function DirectoryIdentityMatch(ByVal name As String, ByVal official As String, ByVal ssn As String) As Long
    Dim db As DAO.Database, r As DAO.Recordset, candidate As Long, matches As Long
    Set db = CurrentDb
    If ssn <> "" Then
        DirectoryIdentityMatch = Nz(MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn) & " AND IsArchived=False"), 0)
        If DirectoryIdentityMatch > 0 Then Exit Function
    End If
    If DirectoryName(name) = "" And DirectoryName(official) = "" Then Exit Function
    Set r = db.OpenRecordset("SELECT MemberID,FullName,ReportName,MemberSSN FROM CommandMember WHERE IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        If (ssn = "" Or Nz(r!MemberSSN, "") = "" Or ssn = Nz(r!MemberSSN, "")) Then
            If (DirectoryName(name) <> "" And DirectoryName(name) = DirectoryName(r!FullName)) Or (DirectoryName(official) <> "" And DirectoryName(official) = DirectoryName(r!ReportName)) Then
                matches = matches + 1: candidate = r!MemberID
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    If matches = 1 Then DirectoryIdentityMatch = candidate
End Function
'Refresh core identity from a confirmed report. Older/undated reports may fill
'blanks but cannot replace information supplied by a later dated report.
Public Sub RefreshDirectoryFromReport(ByVal memberID As Long, ByVal reportID As Long)
    Dim db As DAO.Database, member As DAO.Recordset, report As DAO.Recordset, latest As DAO.Recordset
    Dim ssn As String, rank As String, official As String, display As String, version As Long, fresh As Boolean, changed As Boolean, incoming As String, reported As Variant, reportedText As String
    Set db = CurrentDb
    Set member = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    Set report = db.OpenRecordset("SELECT OfficialName,MemberSSN,RankRate,PeriodEnd,DateReported FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    ssn = Nz(member!MemberSSN, ""): incoming = DirectorySSN(report!MemberSSN)
    If incoming <> "" And ssn <> "" And incoming <> ssn Then member.Close: report.Close: Exit Sub
    'A report's date reported moves the profile's forward, never back.
    reported = Null
    reportedText = Nz(report!DateReported, "")
    If IsValidIsoDate(reportedText) Then
        reported = DateSerial(CLng(Left$(reportedText, 4)), CLng(Mid$(reportedText, 6, 2)), CLng(Right$(reportedText, 2)))
        If Year(reported) < 1950 Or reported > DateAdd("yyyy", 1, Date) Then
            reported = Null
        ElseIf Not IsNull(member!DateReported) Then
            If reported <= member!DateReported Then reported = Null
        End If
    End If
    If (incoming = "" Or incoming = ssn) And (Nz(report!RankRate, "") = "" Or Nz(report!RankRate, "") = Nz(member!RankRate, "")) And (Nz(report!OfficialName, "") = "" Or Nz(report!OfficialName, "") = Nz(member!ReportName, "")) And IsNull(reported) Then member.Close: report.Close: Exit Sub
    Set latest = db.OpenRecordset("SELECT Max(r.PeriodEnd) AS LatestEnd FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND m.MatchState='Confirmed' AND r.ReportID<>" & reportID, dbOpenSnapshot)
    fresh = False
    If Not IsNull(report!PeriodEnd) Then
        If IsNull(latest!LatestEnd) Then fresh = True Else fresh = (report!PeriodEnd >= latest!LatestEnd)
    End If
    latest.Close
    display = member!FullName: official = Nz(member!ReportName, ""): rank = Nz(member!RankRate, ""): version = member!VersionCounter
    If incoming <> "" And ssn = "" Then
        If Nz(MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(incoming) & " AND MemberID<>" & memberID), 0) = 0 Then ssn = incoming: changed = True
    End If
    incoming = Nz(report!RankRate, "")
    If incoming <> "" And incoming <> rank And (fresh Or rank = "") Then rank = incoming: changed = True
    incoming = Nz(report!OfficialName, "")
    If incoming <> "" And incoming <> official And (fresh Or official = "") Then official = incoming: changed = True
    member.Close: report.Close
    If Not IsNull(reported) Then
        SaveDirectoryMember memberID, version, display, official, rank, ssn, "Report", , , reported
    ElseIf changed Then
        SaveDirectoryMember memberID, version, display, official, rank, ssn, "Report"
    End If
End Sub
'P06: a report people field filled from a directory member. Signature names are
'"LAST, F M, RANK" within 28 characters; the counselor is "LAST, F M" within 20, or
'split into last name and initials. Raises when the name cannot fit.
Public Function MemberPeopleText(ByVal memberID As Long, ByVal key As String) As String
    Dim r As DAO.Recordset, brief As String, lastName As String, initials As String, rank As String, limit As Long
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 6970, "modCommandDirectory.MemberPeopleText", "Choose a member from the directory."
    brief = Nz(r!ReportName, ""): If Len(brief) = 0 Then brief = Nz(r!FullName, "")
    rank = Nz(r!RankRate, ""): r.Close
    brief = NavyNameFormat(brief, True)
    If InStr(brief, ",") > 0 Then
        lastName = Trim$(Left$(brief, InStr(brief, ",") - 1)): initials = Trim$(Mid$(brief, InStr(brief, ",") + 1))
    Else
        lastName = brief
    End If
    Select Case key
        Case "RaterName", "SeniorRaterName"
            limit = 28: MemberPeopleText = brief & IIf(Len(rank) > 0, ", " & rank, "")
            If Len(MemberPeopleText) > limit Then MemberPeopleText = lastName & IIf(Len(rank) > 0, ", " & rank, "")
        Case "CounselorName"
            limit = 20: MemberPeopleText = brief
            If Len(MemberPeopleText) > limit Then MemberPeopleText = lastName
        Case "CounselorLastName": limit = 20: MemberPeopleText = lastName
        Case "CounselorFirstInitial": limit = 1: MemberPeopleText = Left$(initials, 1)
        Case "CounselorMiddleInitial": limit = 3: MemberPeopleText = Replace(Trim$(Mid$(initials, 2)), " ", "")
        Case Else: Err.Raise vbObjectError + 6970, "modCommandDirectory.MemberPeopleText", "This field is not filled from the directory."
    End Select
    If Len(MemberPeopleText) > limit Then Err.Raise vbObjectError + 6971, "modCommandDirectory.MemberPeopleText", "The member's name does not fit this field (" & limit & " characters). Type it instead."
End Function
'A new report for a member starts with the profile's date reported (block 9).
Public Sub MemberCreationDefaults(ByVal values As Object, ByVal memberID As Long)
    Dim reported As Variant
    reported = MajorScalar("SELECT DateReported FROM CommandMember WHERE MemberID=" & memberID)
    If Not IsNull(reported) Then values("DateReported") = Format$(reported, "yyyy-mm-dd")
End Sub
Private Function DirectoryName(ByVal value As Variant) As String
    Dim s As String
    s = UCase$(Trim$(Nz(value, "")))
    s = Replace(Replace(Replace(s, ",", ""), ".", ""), " ", "")
    DirectoryName = s
End Function
'A profile without a rank/rate (MNA lists none for officers) takes it from the
'member's latest confirmed report that has one. Returns True when it filled one.
Public Function FillRankFromLatestReport(ByVal memberID As Long) As Boolean
    Dim r As DAO.Recordset, rank As String
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    If Len(Nz(r!RankRate, "")) > 0 Then r.Close: Exit Function
    rank = Trim$(Nz(MajorScalar("SELECT TOP 1 r.RankRate FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & _
        " AND m.MatchState='Confirmed' AND Len(Trim(Nz(r.RankRate,'')))>0 ORDER BY r.PeriodEnd DESC,r.ReportID DESC"), ""))
    If Len(rank) = 0 Or Len(rank) > 5 Then r.Close: Exit Function
    Call SaveDirectoryMember(memberID, r!VersionCounter, r!FullName, Nz(r!ReportName, ""), rank, Nz(r!MemberSSN, ""))
    r.Close: FillRankFromLatestReport = True
End Function
Public Function SaveDirectoryMember(ByVal memberID As Long, ByVal expectedVersion As Long, ByVal displayName As String, ByVal reportName As String, ByVal rank As String, ByVal ssn As String, Optional ByVal sourceKind As String = "Manual", Optional ByVal ruadValues As Variant, Optional ByVal ruadSourceXML As String = "", Optional ByVal dateReported As Variant) As Long
    Dim db As DAO.Database, ruad As Object, ruadXML As String
    Dim r As DAO.Recordset, w As DAO.Workspace, author As Long, failure As Long, message As String, op As Long, duplicate As Variant
    author = MajorAuthor(): displayName = Trim$(displayName): ssn = Trim$(ssn)
    If displayName = "" Or Len(displayName) > 120 Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Enter a member name within 120 characters."
    If Len(reportName) > 27 Or Len(rank) > 5 Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Report name permits 27 characters; rank/rate permits 5. No value will be truncated."
    If ssn <> "" And DirectorySSN(ssn) = "" Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Enter all nine SSN digits or leave SSN unentered."
    'Date reported: left as it is when not passed; blank clears it.
    If Not IsMissing(dateReported) Then
        If Len(Trim$(Nz(dateReported, ""))) = 0 Then
            dateReported = Null
        ElseIf Not IsDate(dateReported) Then
            Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Enter date reported as a date, or leave it blank."
        ElseIf Year(CDate(dateReported)) < 1950 Or CDate(dateReported) > DateAdd("yyyy", 1, Date) Then
            Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Date reported must be a real date, not more than a year ahead."
        Else
            dateReported = DateValue(CDate(dateReported))
        End If
    End If
    If ssn <> "" Then
        duplicate = MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn) & " AND MemberID<>" & memberID)
        If Not IsNull(duplicate) Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "That SSN already belongs to a directory member. Choose the existing member; names alone are not an identity match."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenDynaset)
    If memberID = 0 Then
        r.AddNew: r!MemberGuid = NewIdentity(): r!CreatedBy = author: r!CreatedAt = Now(): r!SourceKind = sourceKind: r!IsArchived = False: r!VersionCounter = 1
    Else
        If r.EOF Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Member no longer exists."
        If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6960, "modCommandDirectory.SaveDirectoryMember", "Member profile changed. Reopen it before saving."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    If IsMissing(ruadValues) Then
        Set ruad = RUADDecode(r!RUADData)
    Else
        Set ruad = ruadValues
    End If
    If ruadSourceXML = "" Then
        ruadXML = RUADEncode(ruad, displayName, r!RUADData)
    Else
        ruadXML = RUADEncode(ruad, displayName, ruadSourceXML)
    End If
    r!RUADData = ruadXML
    r!FullName = displayName: r!ReportName = reportName: r!RankRate = rank
    If ssn = "" Then r!MemberSSN = Null Else r!MemberSSN = ssn
    If Not IsMissing(dateReported) Then r!DateReported = dateReported
    r!ModifiedBy = author: r!ModifiedAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveDirectoryMember = r!MemberID: r.Close
    op = BeginOperation("Save command member", displayName & "; rank/rate " & rank & "; SSN " & IIf(ssn = "", "unentered", "ending " & Right$(ssn, 4)))
    FinishOperation op, "Completed", "Directory profile saved; report fields and prior evaluation records unchanged."
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modCommandDirectory", "SaveDirectoryMember", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function SyncReportDirectory(ByVal reportID As Long) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim report As DAO.Recordset, link As DAO.Recordset, member As DAO.Recordset, id As Long, ssn As String, name As String, official As String, state As String, source As String, existing As Variant
    Set report = ownedDB.OpenRecordset("SELECT MemberName,OfficialName,MemberSSN,RankRate FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If report.EOF Then report.Close: Exit Function
    official = Nz(report!OfficialName, ""): name = UXIdentity(official, report!MemberName)
    If name = "Unnamed member" Then report.Close: Exit Function
    ssn = DirectorySSN(report!MemberSSN)
    Set link = ownedDB.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & reportID, dbOpenDynaset)
    If creationMember > 0 Then
        id = creationMember: state = creationState: source = creationSource
    ElseIf Not link.EOF Then
        id = link!MemberID: state = link!MatchState: source = link!MatchSource
        Set member = ownedDB.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
        If ssn <> "" And ssn <> Nz(member!MemberSSN, "") Then
            state = "Needs review"
            existing = MajorScalar("SELECT MemberID FROM CommandMember WHERE MemberSSN=" & QuoteSQL(ssn))
            If IsNull(existing) And Nz(member!MemberSSN, "") = "" Then
                state = "Confirmed"
            End If
        End If
        member.Close
    Else
        existing = DirectoryIdentityMatch(name, official, ssn)
        If existing = 0 Then
            id = SaveDirectoryMember(0, 0, name, official, Nz(report!RankRate, ""), ssn, "Report")
            state = IIf(ssn = "", "Provisional", "Confirmed"): source = "Report creation"
        Else
            id = CLng(existing): source = "Unique name match": state = "Confirmed"
            Set member = ownedDB.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
            If ssn <> "" And ssn = Nz(member!MemberSSN, "") Then source = "SSN match"
            member.Close
        End If
    End If
    If link.EOF Then link.AddNew: link!ReportID = reportID Else link.Edit
    link!MemberID = id: link!MatchState = state: link!MatchSource = source: link!UpdatedBy = MajorAuthor(): link!UpdatedAt = Now(): link.Update
    link.Close: report.Close
    If state = "Confirmed" Then RefreshDirectoryFromReport id, reportID
    SyncReportDirectory = id
End Function
'Repair only the known PDF placeholder bug; preserve genuine profiles and history.
Public Function RepairPDFDirectoryLinks() As Long
    Dim db As DAO.Database, r As DAO.Recordset, link As DAO.Recordset, w As DAO.Workspace, candidates As New Collection, candidate As Variant, fields As Object, official As String, ssn As String
    Dim oldID As Long, reportID As Long, newID As Long, failure As Long, message As String, stage As String, matched As Boolean, operation As Long
    Set db = CurrentDb
    On Error GoTo Failed
    stage = "Find legacy placeholders"
    Set r = db.OpenRecordset("SELECT DISTINCT m.ReportID,m.MemberID FROM (MemberReport AS m INNER JOIN CommandMember AS c ON m.MemberID=c.MemberID) INNER JOIN ImportReceipt AS i ON m.ReportID=i.ReportID WHERE c.FullName='Imported PDF member' AND c.MemberSSN Is Null AND c.SourceKind='Report' AND c.IsArchived=False AND m.MatchSource='Report creation' AND i.AdapterVersion Like 'pdf-xfa-v*' AND i.IsUndone=False", dbOpenSnapshot)
    Do While Not r.EOF
        candidates.Add Array(CLng(r!ReportID), CLng(r!MemberID)): r.MoveNext
    Loop
    r.Close
    stage = "Begin repair": Set w = DBEngine.Workspaces(0): w.BeginTrans
    For Each candidate In candidates
        reportID = candidate(0): oldID = candidate(1): Set fields = SavedFields(reportID)
        official = Nz(fields("OfficialName"), ""): ssn = DirectorySSN(fields("MemberSSN"))
        If Len(official) > 0 Then
            stage = "Match extracted identity": newID = DirectoryIdentityMatch(official, official, ssn)
            matched = (newID > 0)
            If newID = 0 Then newID = SaveDirectoryMember(0, 0, official, official, Nz(fields("RankRate"), ""), ssn, "Report")
            stage = "Replace legacy link"
            Set link = db.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & reportID, dbOpenDynaset)
            link.Edit: link!MemberID = newID: link!MatchState = IIf(ssn <> "" Or matched, "Confirmed", "Provisional")
            link!MatchSource = "Repaired PDF identity": link!UpdatedBy = MajorAuthor(): link!UpdatedAt = Now(): link.Update: link.Close
            If ssn <> "" Or matched Then RefreshDirectoryFromReport newID, reportID
            stage = "Retire unused placeholder"
            'Use the transaction's own DAO connection; domain aggregates can
            'still see the pre-transaction member link through the UI connection.
            Set link = db.OpenRecordset("SELECT Count(*) AS N FROM MemberReport WHERE MemberID=" & oldID, dbOpenSnapshot)
            If link!n = 0 Then
                link.Close: Set link = db.OpenRecordset("SELECT Count(*) AS N FROM PriorEvaluation WHERE MemberID=" & oldID, dbOpenSnapshot)
                If link!n = 0 Then db.Execute "UPDATE CommandMember SET IsArchived=True WHERE MemberID=" & oldID, dbFailOnError
            End If
            link.Close
            operation = BeginOperation("Repair PDF member identity", "Report " & reportID & "; placeholder member " & oldID & "; matched member " & newID)
            FinishOperation operation, "Completed", "Report values, import receipt and historical evaluations retained."
            RepairPDFDirectoryLinks = RepairPDFDirectoryLinks + 1
        End If
    Next
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modCommandDirectory", "RepairPDFDirectoryLinks", stage, Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = stage & ": " & CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function SyncDirectoryFromReports() As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT ReportID FROM Report ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        If SyncReportDirectory(r!ReportID) > 0 Then SyncDirectoryFromReports = SyncDirectoryFromReports + 1
        r.MoveNext
    Loop
    r.Close
End Function
Public Sub LinkReportDirectoryMember(ByVal reportID As Long, ByVal memberID As Long, ByVal expectedRevision As Long, ByVal expectedMemberVersion As Long, ByVal reason As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim w As DAO.Workspace, r As DAO.Recordset, op As Long, failure As Long, message As String
    If Trim$(reason) = "" Then Err.Raise vbObjectError + 6961, "modCommandDirectory.LinkReportDirectoryMember", "Record why this is the correct member."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If CurrentRevision(reportID) <> expectedRevision Or Nz(MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & memberID), 0) <> expectedMemberVersion Then Err.Raise vbObjectError + 6961, "modCommandDirectory.LinkReportDirectoryMember", "Report or member changed. Preview the identity link again."
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 6961, "modCommandDirectory.LinkReportDirectoryMember", "Choose an active directory member."
    Set r = ownedDB.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & reportID, dbOpenDynaset)
    op = BeginOperation("Confirm report member identity", "Report " & reportID & "; prior directory member " & IIf(r.EOF, "unlinked", Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & reportID), 0)) & "; new member " & memberID & "; reason " & reason)
    If r.EOF Then r.AddNew: r!ReportID = reportID Else r.Edit
    r!MemberID = memberID: r!MatchState = "Confirmed": r!MatchSource = "Manual confirmation": r!UpdatedBy = MajorAuthor(): r!UpdatedAt = Now(): r.Update: r.Close
    FinishOperation op, "Completed", "Identity link confirmed; saved report fields unchanged."
    w.CommitTrans: On Error GoTo 0
    Call FillRankFromLatestReport(memberID)
    Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modCommandDirectory", "LinkReportDirectoryMember", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Sub
Public Sub DirectoryCreationContext(ByVal memberID As Long, Optional ByVal state As String = "Confirmed", Optional ByVal source As String = "Manual selection")
    creationMember = memberID: creationState = state: creationSource = source
End Sub
