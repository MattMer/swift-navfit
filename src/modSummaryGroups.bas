Attribute VB_Name = "modSummaryGroups"
Option Compare Binary
Option Explicit
Public Function CreateSummaryGroup(ByVal title As String, ByVal kind As String, ByVal paygrade As String) As Long
    Dim r As DAO.Recordset, author As Long, allowed As Boolean, item As Variant, grades As Variant, operation As Long, w As DAO.Workspace, failure As Long, message As String
    author = MajorAuthor()
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4100, , "Enter a group title."
    If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 4101, , "Invalid report kind."
    Select Case kind
        Case "EVAL": grades = Array("E1", "E2", "E3", "E4", "E5", "E6")
        Case "CHIEFEVAL": grades = Array("E7", "E8", "E9")
        Case "FITREP": grades = Array("W1", "W2", "W3", "W4", "W5", "O1", "O2", "O3", "O4", "O5", "O6")
    End Select
    For Each item In grades: If item = paygrade Then allowed = True
    Next
    If Not allowed Then Err.Raise vbObjectError + 4102, , "Choose a paygrade appropriate to this report kind."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SummaryGroup", dbOpenDynaset)
    r.AddNew: r!GroupGuid = NewIdentity(): r!Title = title: r!FormKind = kind: r!Paygrade = paygrade: r!VersionCounter = 1: r!IsArchived = False
    r.Update: r.Bookmark = r.LastModified: CreateSummaryGroup = r!GroupID: r.Close
    operation = BeginOperation("Create summary group", "Group " & CreateSummaryGroup & ": " & title & " / " & kind & " / " & paygrade)
    FinishOperation operation, "Completed", "Explicit summary group created. Eligibility still requires review."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "CreateSummaryGroup", message
End Function
Public Function GroupStamp(ByVal groupID As Long) As String
    Dim r As DAO.Recordset, g As DAO.Recordset
    Set g = CurrentDb.OpenRecordset("SELECT * FROM SummaryGroup WHERE GroupID=" & groupID, dbOpenSnapshot)
    If g.EOF Then g.Close: Err.Raise vbObjectError + 4103, , "Group does not exist."
    GroupStamp = g!GroupGuid & ":" & g!VersionCounter & ":" & CStr(g!IsArchived): g.Close
    Set r = CurrentDb.OpenRecordset("SELECT m.ReportID,r.CurrentRevisionID,r.IsArchived,r.WorkflowState FROM GroupMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.GroupID=" & groupID & " ORDER BY m.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        GroupStamp = GroupStamp & "|" & r!ReportID & ":" & Nz(r!CurrentRevisionID, 0) & ":" & CStr(r!IsArchived)
        r.MoveNext
    Loop
    r.Close
End Function
Public Sub AddGroupMember(ByVal groupID As Long, ByVal reportID As Long, ByVal paygrade As String, Optional ByVal exceptionReason As String = "")
    Dim g As DAO.Recordset, r As DAO.Recordset, f As Object, author As Long, w As DAO.Workspace, operation As Long, failure As Long, message As String, db As DAO.Database, stage As String
    On Error GoTo Failed
    stage = "Begin membership transaction": author = MajorAuthor(): Set w = DBEngine.Workspaces(0): w.BeginTrans
    Set db = CurrentDb
    stage = "Read saved member": Set f = SavedFields(reportID)
    Set g = db.OpenRecordset("SELECT * FROM SummaryGroup WHERE GroupID=" & groupID, dbOpenSnapshot)
    stage = "Check group": If g.EOF Then Err.Raise vbObjectError + 4103, , "Group does not exist."
    If g!IsArchived Then Err.Raise vbObjectError + 4104, , "Recover the group first."
    If g!FormKind <> f("FormKind") Or g!Paygrade <> paygrade Then
        Err.Raise vbObjectError + 4105, , "Report kind and explicit paygrade must match the group."
    End If
    g.Close: Set g = Nothing
    stage = "Check member availability": If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 4106, , "Archived report cannot join a group."
    Set r = db.OpenRecordset("GroupMember", dbOpenDynaset)
    stage = "Insert membership": r.AddNew: r!GroupID = groupID: r!ReportID = reportID: r!Paygrade = paygrade: r!ExceptionReason = exceptionReason: r.Update: r.Close: Set r = Nothing
    CurrentDb.Execute "UPDATE SummaryGroup SET VersionCounter=VersionCounter+1 WHERE GroupID=" & groupID, dbFailOnError
    operation = BeginOperation("Add group member", "Group " & groupID & ", report " & reportID & ", explicit paygrade " & paygrade & ": " & exceptionReason)
    FinishOperation operation, "Completed", "Membership added; committed report values unchanged."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not g Is Nothing Then g.Close
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "AddGroupMember", stage & ": " & message
End Sub
Public Sub RemoveGroupMember(ByVal groupID As Long, ByVal reportID As Long)
    Dim author As Long, operation As Long, w As DAO.Workspace, failure As Long, message As String
    author = MajorAuthor(): Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 4104, , "Select an active group."
    If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & reportID) <> 1 Then Err.Raise vbObjectError + 4109, , "Report is not in the group."
    operation = BeginOperation("Remove group member", "Group " & groupID & ", report " & reportID & "; previous exception: " & Nz(MajorScalar("SELECT ExceptionReason FROM GroupMember WHERE GroupID=" & groupID & " AND ReportID=" & reportID), ""))
    CurrentDb.Execute "DELETE FROM GroupMember WHERE GroupID=" & groupID & " AND ReportID=" & reportID, dbFailOnError
    CurrentDb.Execute "UPDATE SummaryGroup SET VersionCounter=VersionCounter+1 WHERE GroupID=" & groupID, dbFailOnError
    FinishOperation operation, "Completed", "Membership removed; frozen rosters retained."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    On Error GoTo 0
    Err.Raise failure, "RemoveGroupMember", message
End Sub
Public Function FreezeSummaryGroup(ByVal groupID As Long, ByVal completenessNote As String) As Long
    Dim w As DAO.Workspace, r As DAO.Recordset, s As DAO.Recordset, m As DAO.Recordset, id As Long, failure As Long, message As String
    If Len(Trim$(completenessNote)) = 0 Then Err.Raise vbObjectError + 4107, , "Record a completeness attestation."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 4104, , "Select an active group."
    Set r = CurrentDb.OpenRecordset("SELECT m.ReportID,r.CurrentRevisionID,r.IsArchived FROM GroupMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.GroupID=" & groupID & " ORDER BY m.ReportID", dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 4108, , "Group is empty."
    Set s = CurrentDb.OpenRecordset("GroupSnapshot", dbOpenDynaset)
    s.AddNew: s!GroupID = groupID: s!Stamp = GroupStamp(groupID): s!AttestedBy = MajorAuthor(): s!AttestedAt = Now(): s!CompletenessNote = completenessNote
    s.Update: s.Bookmark = s.LastModified: id = s!SnapshotID: s.Close
    Set m = CurrentDb.OpenRecordset("GroupSnapshotMember", dbOpenDynaset)
    Do While Not r.EOF
        If r!IsArchived Or Nz(r!CurrentRevisionID, 0) = 0 Then Err.Raise vbObjectError + 4106, , "Group contains an unavailable report."
        m.AddNew: m!SnapshotID = id: m!ReportID = r!ReportID: m!RevisionID = r!CurrentRevisionID
        If MajorRevisionHasField(r!CurrentRevisionID, "PromotionRecommendation") Then m!Promotion = MajorRevisionValue(r!CurrentRevisionID, "PromotionRecommendation") Else m!Promotion = Null
        m!TraitAverage = NumericTraitAverage(r!CurrentRevisionID): m.Update
        r.MoveNext
    Loop
    m.Close: r.Close: w.CommitTrans: FreezeSummaryGroup = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    If Not m Is Nothing Then m.Close
    On Error GoTo 0
    Err.Raise failure, "FreezeSummaryGroup", message
End Function
