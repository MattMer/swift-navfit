Attribute VB_Name = "modReportState"
Option Compare Binary
Option Explicit

Public Function PersistenceCount(ByVal tableName As String, ByVal predicate As String) As Long
    'Use the DAO workspace owning the transaction, including newly imported rows.
    PersistenceCount = CLng(MajorScalar("SELECT Count(*) FROM [" & tableName & "] WHERE " & predicate))
End Function

'ReportHead owns the selected saved revision. Report's value columns and pointer
'are an interchange/query cache; every service read/save checks that projection.
Public Function ReportHeadRevision(ByVal reportID As Long) As Long
    ReportHeadRevision = Nz(MajorScalar("SELECT RevisionID FROM ReportHead WHERE ReportID=" & reportID), 0)
    If ReportHeadRevision = 0 Then Err.Raise vbObjectError + 8900, , "Report has no authoritative saved revision. Inspect persistence diagnostics."
End Function

Private Function StateText(ByVal value As Variant) As Variant
    If IsNull(value) Then
        StateText = Null
    ElseIf VarType(value) = vbDate Then
        StateText = Format$(value, "yyyy-mm-dd")
    Else
        StateText = CStr(value)
    End If
End Function

Public Function RevisionFields(ByVal revisionID As Long) As Object
    Dim values As Object, r As DAO.Recordset, key As Variant
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In DraftFieldKeys(): values.Add CStr(key), Null: Next
    Set r = CurrentDb.OpenRecordset("SELECT FieldKey,ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID, dbOpenSnapshot)
    Do While Not r.EOF
        If values.Exists(CStr(r!FieldKey)) Then
            If r!IsNullValue Then values(CStr(r!FieldKey)) = Null Else values(CStr(r!FieldKey)) = r!ValueText.Value
        End If
        r.MoveNext
    Loop
    r.Close: Set RevisionFields = values
End Function

Public Sub ValidateCatalogueField(ByVal key As String, ByVal value As Variant)
    Dim options As String, item As Variant, valid As Boolean, kind As String
    If IsNull(value) Then Exit Sub
    kind = CatalogueStorageType(key)
    If kind = "DATETIME" Then
        If VarType(value) <> vbDate And Not IsValidIsoDate(CStr(value)) Then Err.Raise vbObjectError + 8910, , key & " requires a real ISO date."
    Else
        If VarType(value) <> vbString Then Err.Raise vbObjectError + 8910, , key & " requires text or Null."
        If UXFieldLimit(key) > 0 And Len(value) > UXFieldLimit(key) Then Err.Raise vbObjectError + 8910, , key & " exceeds its canonical field limit."
        If UXFieldDate(key) And Len(value) > 0 Then
            If Not IsValidIsoDate(CStr(value)) Then Err.Raise vbObjectError + 8910, , key & " requires a real ISO date."
        End If
    End If
    options = UXFieldOptions(key)
    If Len(options) > 0 And Len(CStr(value)) > 0 Then
        For Each item In Split(options, ";"): If item = value Then valid = True
        Next
        If Not valid Then Err.Raise vbObjectError + 8910, , key & " is outside its canonical choice domain."
    End If
End Sub

Public Sub AssertReportCache(ByVal reportID As Long)
    Dim db As DAO.Database, r As DAO.Recordset, v As DAO.Recordset, head As Long, key As Variant, values As Object
    Set db = CurrentDb: head = ReportHeadRevision(reportID)
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 8901, , "Report does not exist."
    If Nz(r!CurrentRevisionID, 0) <> head Then Err.Raise vbObjectError + 8902, , "Derived revision pointer differs from saved state. No values were changed."
    Set values = CreateObject("Scripting.Dictionary")
    Set v = db.OpenRecordset("SELECT FieldKey,ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & head, dbOpenSnapshot)
    Do While Not v.EOF
        If v!IsNullValue Then values(CStr(v!FieldKey)) = Null Else values(CStr(v!FieldKey)) = v!ValueText.Value
        v.MoveNext
    Loop
    v.Close
    For Each key In DraftFieldKeys()
        If values.Exists(CStr(key)) Then
            If Not SameFieldValue(StateText(r.Fields(CStr(key)).Value), values(CStr(key))) Then Err.Raise vbObjectError + 8903, , "Derived report cache differs from revision " & head & " at " & key & ". Preserve both values and review persistence diagnostics."
        ElseIf Nz(MajorScalar("SELECT FieldDefinitionVersion FROM ReportRevision WHERE RevisionID=" & head), 0) >= CatalogueFieldVersion(CStr(key)) Or Not IsNull(r.Fields(CStr(key)).Value) Then
            Err.Raise vbObjectError + 8904, , "Saved revision is missing " & key & ". No values were inferred."
        End If
    Next
    r.Close
    AssertRevisionFacts head, values
End Sub

Public Sub AssertRevisionFacts(ByVal revisionID As Long, ByVal values As Object)
    Dim r As DAO.Recordset, expected As Long, actual As Long, key As Variant, state As String, family As String
    For Each key In Array("DateReported", "CounselingDate", "RaterDate", "SeniorRaterDate", "ReadinessDate")
        If values.Exists(CStr(key)) Then expected = expected + 1
    Next
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RevisionDate WHERE RevisionID=" & revisionID, dbOpenSnapshot)
    Do While Not r.EOF
        key = CStr(r!FieldKey)
        If Not values.Exists(CStr(key)) Then Err.Raise vbObjectError + 8911, , "Typed date has no original snapshot field."
        state = DateFactState(values(key), CStr(key))
        If state <> CStr(r!DateState) Then Err.Raise vbObjectError + 8911, , "Typed date state differs from its original snapshot."
        If state = "Date" Then
            If IsNull(r!DateValue) Then Err.Raise vbObjectError + 8911, , "Typed date is missing."
            If Format$(r!DateValue, "yyyy-mm-dd") <> CStr(values(key)) Then Err.Raise vbObjectError + 8911, , "Typed date differs from its original snapshot."
        End If
        actual = actual + 1: r.MoveNext
    Loop
    r.Close
    If actual <> expected Then Err.Raise vbObjectError + 8911, , "Typed date projection is incomplete."
    expected = 0: actual = 0
    For Each key In Array("Retention", "Statement")
        If values.Exists(CStr(key) & "Yes") And values.Exists(CStr(key) & "No") Then expected = expected + 1
    Next
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RevisionChoice WHERE RevisionID=" & revisionID, dbOpenSnapshot)
    Do While Not r.EOF
        family = CStr(r!ChoiceKind)
        If Not values.Exists(family & "Yes") Or Not values.Exists(family & "No") Then Err.Raise vbObjectError + 8912, , "Typed choice has no original snapshot pair."
        If CStr(r!ChoiceValue) <> PairChoice(values(family & "Yes"), values(family & "No")) Then Err.Raise vbObjectError + 8912, , "Typed choice differs from its original snapshot."
        actual = actual + 1: r.MoveNext
    Loop
    r.Close
    If actual <> expected Then Err.Raise vbObjectError + 8912, , "Typed choice projection is incomplete."
End Sub

Public Function CommitReportChoice(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal family As String, ByVal choice As String) As Long
    Dim changes As Object
    If family <> "Retention" And family <> "Statement" Then Err.Raise vbObjectError + 8913, , "Choose a supported response."
    If choice <> "Yes" And choice <> "No" And choice <> "Unspecified" Then Err.Raise vbObjectError + 8913, , "Choose Yes, No or Unspecified. Conflicting legacy entries require an explicit choice."
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add family & "Yes", IIf(choice = "Yes", "Yes", "No")
    changes.Add family & "No", IIf(choice = "No", "Yes", "No")
    CommitReportChoice = CommitSavedFields(reportID, expectedRevision, changes, "Choose " & family & ": " & choice)
End Function

Public Sub SelectReportHead(ByVal reportID As Long, ByVal revisionID As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "DELETE FROM ReportHead WHERE ReportID=" & reportID, dbFailOnError
    db.Execute "INSERT INTO ReportHead (ReportID,RevisionID) VALUES (" & reportID & "," & revisionID & ")", dbFailOnError
End Sub

Public Function DateFactState(ByVal value As Variant, ByVal key As String) As String
    If IsNull(value) Then
        DateFactState = "Missing"
    ElseIf CStr(value) = "" Then
        DateFactState = "Blank"
    ElseIf key = "CounselingDate" And IsCounselingStatus(CStr(value)) Then
        DateFactState = CStr(value)
    ElseIf IsValidIsoDate(CStr(value)) Then
        DateFactState = "Date"
    Else
        DateFactState = "Unknown"
    End If
End Function

Public Function PairChoice(ByVal yesValue As Variant, ByVal noValue As Variant) As String
    Dim y As String, n As String
    y = Nz(yesValue, ""): n = Nz(noValue, "")
    If (y <> "" And y <> "Yes" And y <> "No") Or (n <> "" And n <> "Yes" And n <> "No") Then
        PairChoice = "Unknown"
    ElseIf y = "Yes" And n = "Yes" Then
        PairChoice = "Conflict"
    ElseIf y = "Yes" Then
        PairChoice = "Yes"
    ElseIf n = "Yes" Then
        PairChoice = "No"
    Else
        PairChoice = "Unspecified"
    End If
End Function

'Typed facts never replace the original snapshot strings, null/empty distinctions,
'or counselor explanation. Unknown/conflicting legacy entries stay reviewable.
Public Sub WriteRevisionFacts(ByVal revisionID As Long)
    Dim db As DAO.Database, r As DAO.Recordset, key As Variant, value As Variant, state As String, family As Variant
    Set db = CurrentDb
    For Each key In Array("DateReported", "CounselingDate", "RaterDate", "SeniorRaterDate", "ReadinessDate")
        If MajorRevisionHasField(revisionID, CStr(key)) And PersistenceCount("RevisionDate", "RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(CStr(key))) = 0 Then
            value = MajorRevisionValue(revisionID, CStr(key)): state = DateFactState(value, CStr(key))
            Set r = db.OpenRecordset("RevisionDate", dbOpenDynaset)
            r.AddNew: r!RevisionID = revisionID: r!FieldKey = key: r!DateState = state
            If state = "Date" Then r!DateValue = DateSerial(CLng(Left$(value, 4)), CLng(Mid$(value, 6, 2)), CLng(Right$(value, 2)))
            r.Update: r.Close
        End If
    Next
    For Each family In Array("Retention", "Statement")
        If MajorRevisionHasField(revisionID, CStr(family) & "Yes") And MajorRevisionHasField(revisionID, CStr(family) & "No") And PersistenceCount("RevisionChoice", "RevisionID=" & revisionID & " AND ChoiceKind=" & QuoteSQL(CStr(family))) = 0 Then
            Set r = db.OpenRecordset("RevisionChoice", dbOpenDynaset)
            r.AddNew: r!RevisionID = revisionID: r!ChoiceKind = family
            r!ChoiceValue = PairChoice(MajorRevisionValue(revisionID, CStr(family) & "Yes"), MajorRevisionValue(revisionID, CStr(family) & "No"))
            r.Update: r.Close
        End If
    Next
End Sub

Public Sub LinkFieldDefault(ByVal defaultID As Long)
    Dim db As DAO.Database, source As DAO.Recordset, link As DAO.Recordset, kind As String, id As Long, valid As Boolean
    Set db = CurrentDb
    Set source = db.OpenRecordset("SELECT * FROM FieldDefault WHERE DefaultID=" & defaultID, dbOpenSnapshot)
    If source.EOF Then Err.Raise vbObjectError + 8905, , "Default does not exist."
    kind = source!ScopeKind: id = source!ScopeID: source.Close
    Set link = db.OpenRecordset("SELECT * FROM FieldDefaultLink WHERE DefaultID=" & defaultID, dbOpenDynaset)
    If link.EOF Then link.AddNew: link!DefaultID = defaultID Else link.Edit
    link!CommandWorkspaceID = Null: link!FolderID = Null: link!GroupID = Null: link!ReportID = Null
    Select Case kind
        Case "Command": valid = (id = 1 And PersistenceCount("Workspace", "WorkspaceID=1") = 1): If valid Then link!CommandWorkspaceID = id
        Case "Folder": valid = (PersistenceCount("Folder", "FolderID=" & id) = 1): If valid Then link!FolderID = id
        Case "Group": valid = (PersistenceCount("SummaryGroup", "GroupID=" & id) = 1): If valid Then link!GroupID = id
        Case "Report": valid = (PersistenceCount("Report", "ReportID=" & id) = 1): If valid Then link!ReportID = id
    End Select
    link!LinkState = IIf(valid, "Verified", "Unresolved"): link.Update: link.Close
End Sub

Public Sub AssertDefaultLink(ByVal defaultID As Long)
    Dim r As DAO.Recordset, linked As Variant
    Set r = CurrentDb.OpenRecordset("SELECT d.ScopeKind,d.ScopeID,l.* FROM FieldDefault AS d INNER JOIN FieldDefaultLink AS l ON d.DefaultID=l.DefaultID WHERE d.DefaultID=" & defaultID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 8906, , "Default has no verified relationship. Review its legacy scope."
    If r!LinkState <> "Verified" Then Err.Raise vbObjectError + 8906, , "Default has an unresolved legacy scope."
    linked = Null
    Select Case CStr(r!ScopeKind)
        Case "Command": linked = r!CommandWorkspaceID.Value
        Case "Folder": linked = r!FolderID.Value
        Case "Group": linked = r!GroupID.Value
        Case "Report": linked = r!ReportID.Value
    End Select
    If IsNull(linked) Then Err.Raise vbObjectError + 8906, , "Default relationship does not match its scope."
    If CLng(linked) <> CLng(r!ScopeID) Then Err.Raise vbObjectError + 8906, , "Default relationship differs from its legacy scope."
    r.Close
End Sub

Public Sub MigrateReportState()
    Dim db As DAO.Database, r As DAO.Recordset, id As Long
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT ReportID,CurrentRevisionID FROM Report", dbOpenSnapshot)
    Do While Not r.EOF
        id = r!ReportID
        If PersistenceCount("ReportHead", "ReportID=" & id) = 0 Then SelectReportHead id, Nz(r!CurrentRevisionID, 0)
        r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT RevisionID FROM ReportRevision", dbOpenSnapshot)
    Do While Not r.EOF: WriteRevisionFacts CLng(r!RevisionID): r.MoveNext: Loop
    r.Close
    Set r = db.OpenRecordset("SELECT ReportID FROM Report", dbOpenSnapshot)
    Do While Not r.EOF: AssertReportCache CLng(r!ReportID): r.MoveNext: Loop
    r.Close
    Set r = db.OpenRecordset("SELECT DefaultID FROM FieldDefault", dbOpenSnapshot)
    Do While Not r.EOF
        If PersistenceCount("FieldDefaultLink", "DefaultID=" & r!DefaultID) = 0 Then LinkFieldDefault CLng(r!DefaultID)
        r.MoveNext
    Loop
    r.Close
End Sub
