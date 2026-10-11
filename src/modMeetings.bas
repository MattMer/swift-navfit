Attribute VB_Name = "modMeetings"
Option Compare Binary
Option Explicit
' Reuse only a matching immutable group baseline with no proposals. Changed group
' stamps or any proposal history require a new set; existing historical rows are not edited.
Public Function ReuseGroupComparison(ByVal groupID As Long, ByVal title As String) As Long
    Dim db As DAO.Database, r As DAO.Recordset, stamp As String, snapshot As Long
    stamp = GroupStamp(groupID)
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT c.SetID,g.Stamp FROM ComparisonSet AS c INNER JOIN GroupSnapshot AS g ON c.GroupSnapshotID=g.SnapshotID WHERE g.GroupID=" & groupID & " AND c.SetID NOT IN (SELECT SetID FROM MeetingProposal) ORDER BY c.SetID DESC", dbOpenSnapshot)
    Do While Not r.EOF
        If StrComp(Nz(r!Stamp, ""), stamp, vbBinaryCompare) = 0 Then ReuseGroupComparison = r!SetID: Exit Do
        r.MoveNext
    Loop
    r.Close
    If ReuseGroupComparison = 0 Then
        snapshot = FreezeSummaryGroup(groupID, "Automatically captured for comparison; membership eligibility is not certified.")
        ReuseGroupComparison = CreateComparisonSet(snapshot, title)
    End If
End Function
' Copy an explicit group snapshot into an ordered comparison set atomically; live
' Members does not call this during ordinary opening or direct editing.
Public Function CreateComparisonSet(ByVal snapshotID As Long, ByVal title As String) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, s As DAO.Recordset, m As DAO.Recordset, id As Long, ordinal As Long, w As DAO.Workspace, failure As Long, message As String
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4400, , "Enter a meeting title."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set m = ownedDB.OpenRecordset("SELECT * FROM GroupSnapshotMember WHERE SnapshotID=" & snapshotID & " ORDER BY ReportID", dbOpenSnapshot)
    If m.EOF Then Err.Raise vbObjectError + 4401, , "Select a populated group snapshot."
    Set r = ownedDB.OpenRecordset("ComparisonSet", dbOpenDynaset)
    r.AddNew: r!GroupSnapshotID = snapshotID: r!Title = title: r!CreatedBy = MajorAuthor(): r!CreatedAt = Now(): r.Update: r.Bookmark = r.LastModified: id = r!SetID: r.Close
    Set s = ownedDB.OpenRecordset("ComparisonMember", dbOpenDynaset)
    Do While Not m.EOF
        ordinal = ordinal + 1: s.AddNew: s!SetID = id: s!ReportID = m!ReportID: s!RevisionID = m!RevisionID: s!SortOrdinal = ordinal: s.Update: m.MoveNext
    Loop
    s.Close: m.Close: w.CommitTrans: CreateComparisonSet = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    If Not m Is Nothing Then m.Close
    On Error GoTo 0
    Err.Raise failure, "CreateComparisonSet", message
End Function
' Record a field proposal against the comparison's saved revision and supersede prior
' proposals for that field; leave the saved report unchanged.
Public Function ProposeMeetingField(ByVal setID As Long, ByVal reportID As Long, ByVal fieldKey As String, ByVal value As Variant) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, revision As Long, fields As Object, author As Long, w As DAO.Workspace, operation As Long, failure As Long, message As String
    author = MajorAuthor()
    revision = Nz(MajorScalar("SELECT RevisionID FROM ComparisonMember WHERE SetID=" & setID & " AND ReportID=" & reportID), 0)
    If revision = 0 Then Err.Raise vbObjectError + 4402, , "Report is not in this meeting snapshot."
    Set fields = SavedFields(reportID)
    If Not fields.Exists(fieldKey) Or fieldKey = "FormKind" Then Err.Raise vbObjectError + 4403, , "Unknown or immutable field."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = ownedDB.OpenRecordset("MeetingProposal", dbOpenDynaset)
    r.AddNew: r!SetID = setID: r!ReportID = reportID: r!BaseRevision = revision: r!FieldKey = fieldKey: r!ValueText = value: r!IsNullValue = IsNull(value): r!ProposalState = "Proposed"
    r.Update: r.Bookmark = r.LastModified: ProposeMeetingField = r!ProposalID: r.Close
    operation = BeginOperation("Meeting field proposal", "Proposal " & ProposeMeetingField & "; meeting " & setID & "; report " & reportID & "; saved revision " & revision & "; field " & fieldKey)
    FinishOperation operation, "Completed", "Proposed is null=" & CStr(IsNull(value)) & "; proposed value=" & Nz(value, "") & "; committed report unchanged."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "ProposeMeetingField", message
End Function
' Apply a single proposal through the same validated atomic batch service.
Public Sub ApplyMeetingProposal(ByVal proposalID As Long)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    ApplyMeetingProposals CStr(proposalID)
End Sub
' Record an order containing every comparison member exactly once without changing report values.
Public Sub SetMeetingOrder(ByVal setID As Long, ByVal reportsCSV As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim ids As Object, key As Variant, ordinal As Long, w As DAO.Workspace, failure As Long, message As String, operation As Long
    Set ids = MajorIDs(reportsCSV)
    If ids.Count <> DCount("*", "ComparisonMember", "SetID=" & setID) Then Err.Raise vbObjectError + 4406, , "Specify every frozen meeting member exactly once."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    For Each key In ids.Keys
        If DCount("*", "ComparisonMember", "SetID=" & setID & " AND ReportID=" & ids(key)) <> 1 Then Err.Raise vbObjectError + 4407, , "Order contains a report outside this meeting."
        ordinal = ordinal + 1
        ownedDB.Execute "UPDATE ComparisonMember SET SortOrdinal=" & ordinal & " WHERE SetID=" & setID & " AND ReportID=" & ids(key), dbFailOnError
    Next
    operation = BeginOperation("Meeting display order", "Meeting " & setID & ": " & reportsCSV)
    FinishOperation operation, "Completed", "Display order recorded; recommendations unchanged."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    On Error GoTo 0
    Err.Raise failure, "SetMeetingOrder", message
End Sub
' Apply selected proposals from one comparison atomically after revision/group/workflow
' checks; record applied revisions and reject stale or already resolved proposals.
Public Sub ApplyMeetingProposals(ByVal proposalIDsCSV As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, changes As Object, w As DAO.Workspace, failure As Long, message As String
    Dim ids As Object, key As Variant, reports As Object, report As Variant, setID As Long, snapshotID As Long, groupID As Long, expected As Object, db As DAO.Database
    Set ids = MajorIDs(proposalIDsCSV): Set reports = CreateObject("Scripting.Dictionary"): Set expected = CreateObject("Scripting.Dictionary")
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set db = CurrentDb
    For Each key In ids.Keys
        Set r = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE ProposalID=" & ids(key), dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 4404, , "Proposal does not exist."
        If r!ProposalState <> "Proposed" Then Err.Raise vbObjectError + 4405, , "Proposal is already resolved."
        If setID = 0 Then setID = r!SetID
        If setID <> r!SetID Then Err.Raise vbObjectError + 4408, , "Select proposals from one meeting."
        report = CStr(r!ReportID)
        If Not reports.Exists(report) Then
            Set changes = CreateObject("Scripting.Dictionary"): reports.Add report, changes: expected.Add report, CLng(r!BaseRevision)
        Else
            Set changes = reports(report)
        End If
        If expected(report) <> r!BaseRevision Then Err.Raise vbObjectError + 4409, , "Proposal revision mismatch."
        If changes.Exists(CStr(r!FieldKey)) Then Err.Raise vbObjectError + 4410, , "Choose one proposal per report field."
        If r!IsNullValue Then changes.Add CStr(r!FieldKey), Null Else changes.Add CStr(r!FieldKey), r!ValueText.Value
        r.Close
    Next
    snapshotID = Nz(MajorScalar("SELECT GroupSnapshotID FROM ComparisonSet WHERE SetID=" & setID), 0)
    groupID = Nz(MajorScalar("SELECT GroupID FROM GroupSnapshot WHERE SnapshotID=" & snapshotID), 0)
    If StrComp(Nz(MajorScalar("SELECT Stamp FROM GroupSnapshot WHERE SnapshotID=" & snapshotID), ""), GroupStamp(groupID), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 4411, , "Group changed since the meeting snapshot. Refresh the comparison and review the saved values."
    For Each report In reports.Keys
        Call CommitSavedFields(CLng(report), CLng(expected(report)), reports(report), "Reviewed meeting " & setID & " proposal batch")
    Next
    For Each key In ids.Keys
        Set r = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE ProposalID=" & ids(key), dbOpenDynaset)
        r.Edit: r!AppliedRevision = CurrentRevision(r!ReportID): r!ProposalState = "Applied": r.Update: r.Close
    Next
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "ApplyMeetingProposals", message
End Sub
