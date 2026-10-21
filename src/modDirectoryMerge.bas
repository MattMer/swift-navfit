Attribute VB_Name = "modDirectoryMerge"
Option Compare Binary
Option Explicit

'Candidate signal only. ReportName is Last, First; display names are First Last.
Public Function DirectorySurname(ByVal fullName As Variant, ByVal reportName As Variant) As String
    Dim s As String, pos As Long, re As Object
    s = Trim$(Nz(reportName, "")): pos = InStr(s, ",")
    If pos > 0 Then
        s = Left$(s, pos - 1)
    Else
        s = Trim$(Nz(fullName, "")): pos = InStrRev(s, " ")
        If pos > 0 Then s = Mid$(s, pos + 1)
    End If
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.Pattern = "[ .'-]"
    DirectorySurname = UCase$(re.Replace(s, ""))
End Function

Public Function DirectoryMergeStamp(ByVal sourceID As Long, ByVal targetID As Long) As String
    Dim db As DAO.Database, r As DAO.Recordset, field As DAO.Field, text As String, query As Variant, v As String, ids As String
    If sourceID <= 0 Or targetID <= 0 Or sourceID = targetID Then Err.Raise vbObjectError + 9100, "DirectoryMergeStamp", "Select two different members."
    Set db = CurrentDb: ids = sourceID & "," & targetID
    For Each query In Array("SELECT * FROM CommandMember WHERE MemberID IN (" & ids & ") ORDER BY MemberID", _
        "SELECT m.*,r.CurrentRevisionID FROM MemberReport AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.MemberID IN (" & ids & ") ORDER BY m.ReportID", _
        "SELECT * FROM PriorEvaluation WHERE MemberID IN (" & ids & ") ORDER BY PriorID", _
        "SELECT t.* FROM PriorEvaluationTrait AS t INNER JOIN PriorEvaluation AS p ON t.PriorID=p.PriorID WHERE p.MemberID IN (" & ids & ") ORDER BY t.PriorID,t.BlockNumber", _
        "SELECT r.ReportID,r.MemberSSN,r.MemberDoDID FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID IN (" & ids & ") ORDER BY r.ReportID")
        text = text & "TABLE;": Set r = db.OpenRecordset(CStr(query), dbOpenSnapshot)
        Do While Not r.EOF
            For Each field In r.Fields
                If IsNull(field.Value) Then
                    text = text & "N;"
                Else
                    If field.Type = dbDate Then v = CStr(CDbl(field.Value)) Else v = CStr(field.Value)
                    text = text & Len(v) & ":" & v & ";"
                End If
            Next
            text = text & "ROW;": r.MoveNext
        Loop
        r.Close
    Next
    DirectoryMergeStamp = SourceTextHash(text)
End Function

Private Sub CheckMergeIdentity(ByVal db As DAO.Database, ByVal sourceID As Long, ByVal targetID As Long)
    Dim r As DAO.Recordset, ids As String, ssn As String, dod As String, value As String, field As Variant
    ids = sourceID & "," & targetID
    Set r = db.OpenRecordset("SELECT MemberID,MemberSSN,IsArchived FROM CommandMember WHERE MemberID IN (" & ids & ")", dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 9101, "CheckMergeIdentity", "Both members must exist."
    If r.RecordCount <> 2 Then
        r.MoveLast
        If r.RecordCount <> 2 Then Err.Raise vbObjectError + 9101, "CheckMergeIdentity", "Both members must exist."
        r.MoveFirst
    End If
    Do While Not r.EOF
        If Nz(r!IsArchived, False) Then Err.Raise vbObjectError + 9102, "CheckMergeIdentity", "Choose two active members."
        value = Trim$(Nz(r!MemberSSN, ""))
        If Len(value) > 0 Then
            If Len(ssn) > 0 And value <> ssn Then Err.Raise vbObjectError + 9103, "CheckMergeIdentity", "Conflicting SSNs. Investigate and correct identity separately; no merge performed."
            ssn = value
        End If
        r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT r.MemberSSN,r.MemberDoDID FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID IN (" & ids & ")", dbOpenSnapshot)
    Do While Not r.EOF
        value = Trim$(Nz(r!MemberSSN, ""))
        If Len(value) > 0 Then
            If Len(ssn) > 0 And ssn <> value Then Err.Raise vbObjectError + 9104, "CheckMergeIdentity", "A linked report has a conflicting SSN. Investigate before merging."
            ssn = value
        End If
        value = Trim$(Nz(r!MemberDoDID, ""))
        If Len(value) > 0 Then
            If Len(dod) > 0 And dod <> value Then Err.Raise vbObjectError + 9105, "CheckMergeIdentity", "Linked reports have conflicting DoD IDs. Investigate before merging."
            dod = value
        End If
        r.MoveNext
    Loop
    r.Close
    If Len(Nz(MajorScalar("SELECT MemberSSN FROM CommandMember WHERE MemberID=" & sourceID), "")) > 0 And Len(Nz(MajorScalar("SELECT MemberSSN FROM CommandMember WHERE MemberID=" & targetID), "")) = 0 Then Err.Raise vbObjectError + 9106, "CheckMergeIdentity", "Keep the member with the entered SSN as the survivor. Swap the selections and preview again."
    Set r = db.OpenRecordset("SELECT TOP 1 MergeID FROM DirectoryMerge WHERE MergeState='Committed' AND (FromMemberID IN (" & ids & ") OR ToMemberID IN (" & ids & "))", dbOpenSnapshot)
    If Not r.EOF Then Err.Raise vbObjectError + 9107, "CheckMergeIdentity", "A selected member already participates in a retained merge. Undo that merge first, or review the separate profile without another merge."
    r.Close
End Sub

Public Function PreviewDirectoryMerge(ByVal sourceID As Long, ByVal targetID As Long) As String
    Dim db As DAO.Database
    Set db = CurrentDb
    If sourceID <= 0 Or targetID <= 0 Or sourceID = targetID Then Err.Raise vbObjectError + 9108, "PreviewDirectoryMerge", "Select two different members."
    CheckMergeIdentity db, sourceID, targetID
    PreviewDirectoryMerge = DirectoryMergeStamp(sourceID, targetID)
End Function

Private Sub MergeEditorGuard()
    Dim name As Variant
    For Each name In Array("frmEditor", "frmUXMemberEditor", "frmUXPriorEditor")
        'AllForms names are enumerated: older UI generations may lack a prior editor.
        Dim form As Access.AccessObject
        For Each form In CurrentProject.AllForms
            If form.Name = name And form.IsLoaded Then Err.Raise vbObjectError + 9109, "MergeEditorGuard", "Close report, member and prior-evaluation editors before changing directory links. Pending entries remain in their editor."
        Next
    Next
End Sub

Public Function ApplyDirectoryMerge(ByVal sourceID As Long, ByVal targetID As Long, ByVal expectedStamp As String, ByVal reason As String) As Long
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, parent As DAO.Recordset
    Dim trans As Boolean, actor As Long, operation As Long, id As Long, failure As Long, message As String, re As Object
    On Error GoTo Failed
    Call MergeEditorGuard: actor = MajorAuthor(): reason = Trim$(reason)
    If Len(reason) = 0 Or Len(reason) > 255 Then Err.Raise vbObjectError + 9110, "ApplyDirectoryMerge", "Enter a review reason of 1 to 255 characters. Do not enter identifiers."
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "[0-9]{9,}": re.Global = True
    If re.Test(Replace(Replace(Replace(reason, "-", ""), " ", ""), ".", "")) Then Err.Raise vbObjectError + 9111, "ApplyDirectoryMerge", "Keep SSNs and DoD IDs out of the review reason."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    db.Execute "UPDATE CommandMember SET VersionCounter=VersionCounter WHERE MemberID IN (" & sourceID & "," & targetID & ")", dbFailOnError
    If Len(expectedStamp) <> 64 Or PreviewDirectoryMerge(sourceID, targetID) <> expectedStamp Then Err.Raise vbObjectError + 9112, "ApplyDirectoryMerge", "The preview is stale. Preview the members and their links again."
    operation = BeginOperation("Directory merge", "Member IDs " & sourceID & " -> " & targetID)
    Set parent = db.OpenRecordset("DirectoryMerge", dbOpenDynaset)
    parent.AddNew: parent!MergeGuid = NewIdentity(): parent!FromMemberID = sourceID: parent!ToMemberID = targetID
    parent!OperationID = operation: parent!MergeState = "Committed": parent!Reason = reason: parent!CreatedBy = actor: parent!CreatedAt = Now()
    Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & sourceID, dbOpenSnapshot)
    parent!SourceWasArchived = r!IsArchived: parent!SourceModifiedBy = r!ModifiedBy: parent!SourceModifiedAt = r!ModifiedAt: r.Close
    Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & targetID, dbOpenSnapshot)
    parent!TargetModifiedBy = r!ModifiedBy: parent!TargetModifiedAt = r!ModifiedAt: r.Close
    parent.Update: parent.Bookmark = parent.LastModified: id = parent!MergeID
    db.Execute "INSERT INTO DirectoryMergeReport (MergeID,ReportID,MatchState,MatchSource,UpdatedBy,UpdatedAt) SELECT " & id & ",ReportID,MatchState,MatchSource,UpdatedBy,UpdatedAt FROM MemberReport WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "INSERT INTO DirectoryMergePrior (MergeID,PriorID,ModifiedBy,ModifiedAt) SELECT " & id & ",PriorID,ModifiedBy,ModifiedAt FROM PriorEvaluation WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "UPDATE MemberReport SET MemberID=" & targetID & ",MatchSource='Reviewed directory merge',UpdatedBy=" & actor & ",UpdatedAt=Now() WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "UPDATE PriorEvaluation SET MemberID=" & targetID & ",VersionCounter=VersionCounter+1,ModifiedBy=" & actor & ",ModifiedAt=Now() WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "UPDATE CommandMember SET IsArchived=True WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "UPDATE CommandMember SET VersionCounter=VersionCounter+1,ModifiedBy=" & actor & ",ModifiedAt=Now() WHERE MemberID IN (" & sourceID & "," & targetID & ")", dbFailOnError
    parent.Edit: parent!AfterFingerprint = DirectoryMergeStamp(sourceID, targetID): parent.Update: parent.Close
    FinishOperation operation, "Completed", "Directory links moved. Original profile, report values, revisions and prior traits retained. Merge ID " & id
    w.CommitTrans: trans = False: ApplyDirectoryMerge = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    Resume CleanupFailure
CleanupFailure:
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not parent Is Nothing Then parent.Close
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "ApplyDirectoryMerge", message
End Function

Public Function UndoDirectoryMerge(ByVal mergeID As Long) As Boolean
    Dim db As DAO.Database, w As DAO.Workspace, parent As DAO.Recordset, rows As DAO.Recordset, item As DAO.Recordset
    Dim sourceID As Long, targetID As Long, trans As Boolean, actor As Long, operation As Long, field As Variant, failure As Long, message As String
    On Error GoTo Failed
    Call MergeEditorGuard: actor = MajorAuthor(): Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
    w.BeginTrans: trans = True
    Set parent = db.OpenRecordset("SELECT * FROM DirectoryMerge WHERE MergeID=" & mergeID, dbOpenDynaset)
    If parent.EOF Then Err.Raise vbObjectError + 9113, "UndoDirectoryMerge", "Select a saved merge."
    If parent!MergeState <> "Committed" Then Err.Raise vbObjectError + 9114, "UndoDirectoryMerge", "This merge has already been undone."
    sourceID = parent!FromMemberID: targetID = parent!ToMemberID
    db.Execute "UPDATE CommandMember SET VersionCounter=VersionCounter WHERE MemberID IN (" & sourceID & "," & targetID & ")", dbFailOnError
    If DirectoryMergeStamp(sourceID, targetID) <> parent!AfterFingerprint Then Err.Raise vbObjectError + 9115, "UndoDirectoryMerge", "Profiles, report revisions, links or prior evaluations changed after merging. Automatic undo is blocked to preserve those changes."
    operation = BeginOperation("Undo directory merge", "Merge ID " & mergeID)
    Set rows = db.OpenRecordset("SELECT * FROM DirectoryMergeReport WHERE MergeID=" & mergeID, dbOpenSnapshot)
    Do While Not rows.EOF
        Set item = db.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & rows!ReportID, dbOpenDynaset)
        item.Edit: item!MemberID = sourceID
        For Each field In Array("MatchState", "MatchSource", "UpdatedBy", "UpdatedAt"): item.Fields(field) = rows.Fields(field): Next
        item.Update: item.Close: rows.MoveNext
    Loop
    rows.Close
    Set rows = db.OpenRecordset("SELECT * FROM DirectoryMergePrior WHERE MergeID=" & mergeID, dbOpenSnapshot)
    Do While Not rows.EOF
        Set item = db.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & rows!PriorID, dbOpenDynaset)
        item.Edit: item!MemberID = sourceID: item!VersionCounter = item!VersionCounter + 1
        item!ModifiedBy = rows!ModifiedBy: item!ModifiedAt = rows!ModifiedAt: item.Update: item.Close: rows.MoveNext
    Loop
    rows.Close
    db.Execute "UPDATE CommandMember SET VersionCounter=VersionCounter+1 WHERE MemberID=" & sourceID, dbFailOnError
    db.Execute "UPDATE CommandMember SET VersionCounter=VersionCounter+1 WHERE MemberID=" & targetID, dbFailOnError
    For Each field In Array(sourceID, targetID)
        Set item = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & field, dbOpenDynaset): item.Edit
        If CLng(field) = sourceID Then
            item!IsArchived = parent!SourceWasArchived: item!ModifiedBy = parent!SourceModifiedBy: item!ModifiedAt = parent!SourceModifiedAt
        Else
            item!ModifiedBy = parent!TargetModifiedBy: item!ModifiedAt = parent!TargetModifiedAt
        End If
        item.Update: item.Close
    Next
    parent.Edit: parent!MergeState = "Undone": parent!UndoneBy = actor: parent!UndoneAt = Now(): parent.Update: parent.Close
    FinishOperation operation, "Completed", "Restored original directory links for merge ID " & mergeID & "; retained merge audit."
    w.CommitTrans: trans = False: UndoDirectoryMerge = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    Resume CleanupFailure
CleanupFailure:
    On Error Resume Next
    If Not item Is Nothing Then item.Close
    If Not rows Is Nothing Then rows.Close
    If Not parent Is Nothing Then parent.Close
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "UndoDirectoryMerge", message
End Function
