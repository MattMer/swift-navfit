Attribute VB_Name = "modOfficialBatch"
Option Compare Binary
Option Explicit
Public Function PreviewReportChanges(ByVal reportsCSV As String, ByVal changes As Object, ByVal note As String) As Long
    Dim ids As Object, reportKey As Variant, fieldKey As Variant, values As Object, w As DAO.Workspace
    Dim db As DAO.Database, b As DAO.Recordset, item As DAO.Recordset, id As Long, failure As Long, message As String
    If Len(Trim$(note)) = 0 Or changes.Count = 0 Then Err.Raise vbObjectError + 4220, , "Change a report block and enter a reason before previewing."
    Set ids = MajorIDs(reportsCSV): Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set b = db.OpenRecordset("ChangeBatch", dbOpenDynaset)
    b.AddNew: b!BatchGuid = NewIdentity(): b!BatchKind = "ReportBlocks": b!BatchState = "Preview": b!AuthorID = MajorAuthor(): b!CreatedAt = Now(): b!BatchNote = note
    b.Update: b.Bookmark = b.LastModified: id = b!BatchID: b.Close: Set b = Nothing
    Set item = db.OpenRecordset("ChangeItem", dbOpenDynaset)
    For Each reportKey In ids.Keys
        Set values = SavedFields(ids(reportKey))
        For Each fieldKey In changes.Keys
            If Not values.Exists(CStr(fieldKey)) Or fieldKey = "FormKind" Then Err.Raise vbObjectError + 4221, , "Unknown or immutable report block."
            item.AddNew: item!BatchID = id: item!ReportID = ids(reportKey): item!ExpectedRevision = CurrentRevision(ids(reportKey)): item!FieldKey = fieldKey
            If VarType(values(fieldKey)) = vbDate Then item!BeforeValue = Format$(values(fieldKey), "yyyy-mm-dd") Else item!BeforeValue = values(fieldKey)
            item!BeforeIsNull = IsNull(values(fieldKey))
            If VarType(changes(fieldKey)) = vbDate Then item!NewValue = Format$(changes(fieldKey), "yyyy-mm-dd") Else item!NewValue = changes(fieldKey)
            item!NewIsNull = IsNull(changes(fieldKey)): item.Update
        Next
    Next
    item.Close: w.CommitTrans: PreviewReportChanges = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not item Is Nothing Then item.Close
    If Not b Is Nothing Then b.Close
    On Error GoTo 0
    Err.Raise failure, "PreviewReportChanges", message
End Function
Public Function PreviewFieldBatch(ByVal reportsCSV As String, ByVal fieldKey As String, ByVal newValue As Variant, ByVal note As String) As Long
    Dim ids As Object, key As Variant, fields As Object, b As DAO.Recordset, r As DAO.Recordset, id As Long, w As DAO.Workspace, failure As Long, message As String
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4200, , "Record the reason for this batch."
    Set ids = MajorIDs(reportsCSV): Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set b = CurrentDb.OpenRecordset("ChangeBatch", dbOpenDynaset)
    b.AddNew: b!BatchGuid = NewIdentity(): b!BatchKind = "FieldChange": b!BatchState = "Preview": b!AuthorID = MajorAuthor(): b!CreatedAt = Now(): b!BatchNote = note
    b.Update: b.Bookmark = b.LastModified: id = b!BatchID: b.Close
    Set r = CurrentDb.OpenRecordset("ChangeItem", dbOpenDynaset)
    For Each key In ids.Keys
        Set fields = SavedFields(ids(key))
        If Not fields.Exists(fieldKey) Or fieldKey = "FormKind" Then Err.Raise vbObjectError + 4201, , "Unknown or immutable field."
        r.AddNew: r!BatchID = id: r!ReportID = ids(key): r!ExpectedRevision = CurrentRevision(ids(key)): r!FieldKey = fieldKey
        If VarType(newValue) = vbDate And (fieldKey = "PeriodStart" Or fieldKey = "PeriodEnd") Then r!NewValue = Format$(newValue, "yyyy-mm-dd") Else r!NewValue = newValue
        r!NewIsNull = IsNull(newValue)
        If VarType(fields(fieldKey)) = vbDate Then r!BeforeValue = Format$(fields(fieldKey), "yyyy-mm-dd") Else r!BeforeValue = fields(fieldKey)
        r!BeforeIsNull = IsNull(fields(fieldKey)): r.Update
    Next
    r.Close: w.CommitTrans: PreviewFieldBatch = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not b Is Nothing Then b.Close
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "PreviewFieldBatch", message
End Function
Public Sub ApplyFieldBatch(ByVal batchID As Long)
    Dim w As DAO.Workspace, b As DAO.Recordset, r As DAO.Recordset, changes As Object, id As Long, rev As Long, expected As Long, failure As Long, message As String
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set b = CurrentDb.OpenRecordset("SELECT * FROM ChangeBatch WHERE BatchID=" & batchID, dbOpenDynaset)
    If b.EOF Then Err.Raise vbObjectError + 4202, , "Batch does not exist."
    If b!BatchState <> "Preview" Or b!AuthorID <> MajorAuthor() Then Err.Raise vbObjectError + 4203, , "Only the preview author can apply a pending batch."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ChangeItem WHERE BatchID=" & batchID & " ORDER BY ReportID,ItemID", dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 4204, , "Batch is empty."
    Do While Not r.EOF
        id = r!ReportID: expected = r!ExpectedRevision: Set changes = CreateObject("Scripting.Dictionary")
        Do While Not r.EOF
            If r!ReportID <> id Then Exit Do
            If r!ExpectedRevision <> expected Then Err.Raise vbObjectError + 4205, , "Inconsistent preview revision."
            If r!NewIsNull Then changes.Add CStr(r!FieldKey), Null Else changes.Add CStr(r!FieldKey), r!NewValue.Value
            r.MoveNext
        Loop
        Call CommitSavedFields(id, expected, changes, "Batch " & batchID & ": " & Left$(Nz(b!BatchNote, ""), 200))
        rev = CurrentRevision(id)
        CurrentDb.Execute "UPDATE ChangeItem SET AppliedRevision=" & rev & " WHERE BatchID=" & batchID & " AND ReportID=" & id, dbFailOnError
    Loop
    r.Close: b.Edit: b!BatchState = "Applied": b.Update: b.Close: w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not b Is Nothing Then b.Close
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "ApplyFieldBatch", message
End Sub
Public Sub SetFieldDefault(ByVal scopeKind As String, ByVal scopeID As Long, ByVal fieldKey As String, ByVal value As Variant)
    Dim key As Variant, known As Boolean, r As DAO.Recordset, author As Long, w As DAO.Workspace, operation As Long, detail As String, failure As Long, message As String
    author = MajorAuthor()
    For Each key In DraftFieldKeys(): If key = fieldKey And key <> "FormKind" Then known = True
    Next
    If Not known Then Err.Raise vbObjectError + 4210, , "Unknown or immutable field."
    Select Case scopeKind
        Case "Command"
            If scopeID <> 1 Then Err.Raise vbObjectError + 4211, , "Invalid command scope."
        Case "Folder"
            If DCount("*", "Folder", "FolderID=" & scopeID) <> 1 Then Err.Raise vbObjectError + 4211, , "Folder does not exist."
        Case "Group"
            If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & scopeID), True) Then Err.Raise vbObjectError + 4211, , "Recover this group before changing defaults."
            If DCount("*", "SummaryGroup", "GroupID=" & scopeID) <> 1 Then Err.Raise vbObjectError + 4211, , "Group does not exist."
        Case "Report"
            If DCount("*", "Report", "ReportID=" & scopeID) <> 1 Then Err.Raise vbObjectError + 4211, , "Report does not exist."
        Case Else: Err.Raise vbObjectError + 4211, , "Invalid default scope."
    End Select
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FieldDefault WHERE ScopeKind=" & QuoteSQL(scopeKind) & " AND ScopeID=" & scopeID & " AND FieldKey=" & QuoteSQL(fieldKey), dbOpenDynaset)
    detail = scopeKind & " " & scopeID & ", field " & fieldKey
    If r.EOF Then
        detail = detail & "; previous default absent"
        r.AddNew: r!ScopeKind = scopeKind: r!ScopeID = scopeID: r!FieldKey = fieldKey: r!VersionCounter = 1
    Else
        detail = detail & "; previous version " & r!VersionCounter & "; previous is null=" & CStr(r!IsNullValue) & "; previous value=" & Nz(r!ValueText, "")
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!ValueText = value: r!IsNullValue = IsNull(value): r.Update: r.Close
    operation = BeginOperation("Scoped default change", detail)
    FinishOperation operation, "Completed", "New is null=" & CStr(IsNull(value)) & "; new value=" & Nz(value, "") & "; committed reports unchanged."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "SetFieldDefault", message
End Sub
Private Function DefaultAt(ByVal scopeKind As String, ByVal scopeID As Long, ByVal fieldKey As String, ByVal result As Object) As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FieldDefault WHERE ScopeKind=" & QuoteSQL(scopeKind) & " AND ScopeID=" & scopeID & " AND FieldKey=" & QuoteSQL(fieldKey), dbOpenSnapshot)
    If Not r.EOF Then
        If r!IsNullValue Then result("Value") = Null Else result("Value") = r!ValueText.Value
        result("Source") = scopeKind & " " & scopeID & " version " & r!VersionCounter: DefaultAt = True
    End If
    r.Close
End Function
Public Function ResolveFieldDefault(ByVal reportID As Long, ByVal fieldKey As String, Optional ByVal groupID As Long = 0) As Object
    Dim result As Object, folderID As Long, ancestors As Collection, id As Variant, seen As Object, i As Long
    Set result = CreateObject("Scripting.Dictionary"): result.Add "Value", Null: result.Add "Source", "No default"
    Set seen = CreateObject("Scripting.Dictionary"): Set ancestors = New Collection
    folderID = Nz(MajorScalar("SELECT FolderID FROM Report WHERE ReportID=" & reportID), 0)
    Do While folderID > 0
        If seen.Exists(CStr(folderID)) Then Err.Raise vbObjectError + 4212, , "Folder hierarchy contains a cycle."
        seen.Add CStr(folderID), True: ancestors.Add folderID
        folderID = Nz(MajorScalar("SELECT ParentFolderID FROM Folder WHERE FolderID=" & folderID), 0)
    Loop
    Call DefaultAt("Command", 1, fieldKey, result)
    For i = ancestors.Count To 1 Step -1: Call DefaultAt("Folder", CLng(ancestors(i)), fieldKey, result): Next
    If groupID > 0 Then
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & reportID) <> 1 Then Err.Raise vbObjectError + 4213, , "Report is not in the selected group."
        Call DefaultAt("Group", groupID, fieldKey, result)
    End If
    Call DefaultAt("Report", reportID, fieldKey, result)
    Set ResolveFieldDefault = result
End Function
Public Function PreviewBatchUndo(ByVal batchID As Long) As Long
    Dim b As DAO.Recordset, r As DAO.Recordset, s As DAO.Recordset, w As DAO.Workspace, id As Long, failure As Long, message As String
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT BatchState FROM ChangeBatch WHERE BatchID=" & batchID), "") <> "Applied" Then Err.Raise vbObjectError + 4214, , "Select an applied batch."
    Set b = CurrentDb.OpenRecordset("ChangeBatch", dbOpenDynaset)
    b.AddNew: b!BatchGuid = NewIdentity(): b!BatchKind = "Undo": b!BatchState = "Preview": b!AuthorID = MajorAuthor(): b!CreatedAt = Now(): b!BatchNote = "Undo batch " & batchID
    b.Update: b.Bookmark = b.LastModified: id = b!BatchID: b.Close
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ChangeItem WHERE BatchID=" & batchID, dbOpenSnapshot)
    Set s = CurrentDb.OpenRecordset("ChangeItem", dbOpenDynaset)
    Do While Not r.EOF
        If CurrentRevision(r!ReportID) <> Nz(r!AppliedRevision, 0) Then Err.Raise vbObjectError + 4215, , "A report changed after the batch; undo cannot replace later work."
        s.AddNew: s!BatchID = id: s!ReportID = r!ReportID: s!ExpectedRevision = r!AppliedRevision: s!FieldKey = r!FieldKey
        s!NewValue = r!BeforeValue: s!NewIsNull = r!BeforeIsNull: s!BeforeValue = r!NewValue: s!BeforeIsNull = r!NewIsNull: s.Update: r.MoveNext
    Loop
    s.Close: r.Close: w.CommitTrans: PreviewBatchUndo = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not b Is Nothing Then b.Close
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    On Error GoTo 0
    Err.Raise failure, "PreviewBatchUndo", message
End Function
