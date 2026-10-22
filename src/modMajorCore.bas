Attribute VB_Name = "modMajorCore"
Option Compare Binary
Option Explicit
Public Function MajorScalar(ByVal sql As String) As Variant
    Dim r As DAO.Recordset, db As DAO.Database
    Set db = CurrentDb: Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    If r.EOF Then MajorScalar = Null Else MajorScalar = r.Fields(0).Value
    r.Close
End Function
Public Function MajorAuthor() As Long
    MajorAuthor = ActiveContributorID()
End Function
Public Function QuoteSQL(ByVal value As String) As String
    QuoteSQL = "'" & Replace(value, "'", "''") & "'"
End Function
Public Function MajorIDs(ByVal csv As String) As Object
    Dim result As Object, item As Variant, i As Long, value As String, n As Long
    Set result = CreateObject("Scripting.Dictionary")
    If Len(Trim$(csv)) = 0 Then Err.Raise vbObjectError + 4001, , "Select at least one report."
    For Each item In Split(csv, ",")
        value = Trim$(CStr(item))
        If Len(value) = 0 Or Len(value) > 9 Then Err.Raise vbObjectError + 4001, , "Invalid report selection."
        For i = 1 To Len(value)
            If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Err.Raise vbObjectError + 4001, , "Invalid report selection."
        Next
        n = CLng(value)
        If n < 1 Then Err.Raise vbObjectError + 4001, , "Invalid report selection."
        If result.Exists(CStr(n)) Then Err.Raise vbObjectError + 4001, , "Duplicate report selection."
        result.Add CStr(n), n
    Next
    Set MajorIDs = result
End Function
Public Function SavedFields(ByVal reportID As Long) As Object
    Dim r As DAO.Recordset, values As Object, key As Variant, db As DAO.Database
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 4002, , "Report does not exist."
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In DraftFieldKeys(): values.Add CStr(key), r.Fields(key).Value: Next
    r.Close
    values("CommandEmployment") = StoredBlock28(reportID)
    Set SavedFields = values
End Function
Public Function FieldsSubset(ByVal values As Object, ByVal keys As Variant) As Variant
    Dim result() As Variant, i As Long
    ReDim result(0 To UBound(keys))
    For i = 0 To UBound(keys): result(i) = values(CStr(keys(i))): Next
    FieldsSubset = result
End Function
Public Function CommitSavedFields(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal changes As Object, ByVal reason As String) As Long
    Dim db As DAO.Database, r As DAO.Recordset, values As Object, key As Variant, w As DAO.Workspace, failure As Long, message As String, stage As String
    On Error GoTo Failed
    stage = "Begin guarded transaction"
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    stage = "Revision guard": Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT CurrentRevisionID,IsArchived,WorkflowState FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 4002, , "Report does not exist."
    If r!IsArchived Or r!WorkflowState <> "Draft" Then r.Close: Err.Raise vbObjectError + 4003, , "Archived or locked report cannot be changed."
    If expectedRevision = 0 Or Nz(r!CurrentRevisionID, 0) <> expectedRevision Then r.Close: Err.Raise vbObjectError + 4004, , "Report changed since preview."
    r.Close: stage = "Read saved fields": Set values = SavedFields(reportID)
    For Each key In changes.Keys
        If Not values.Exists(key) Or key = "FormKind" Then Err.Raise vbObjectError + 4005, , "Unknown or immutable draft field."
        If (key = "Title" Or key = "MemberName" Or key = "DraftNotes") And IsNull(changes(key)) Then Err.Raise vbObjectError + 4008, , "This draft field uses empty text. Clear the null selection and use an empty value."
        values(key) = changes(key)
        If key = "PeriodStart" Or key = "PeriodEnd" Then
            If Not IsNull(values(key)) And VarType(values(key)) <> vbDate Then
                If Not IsValidIsoDate(CStr(values(key))) Then Err.Raise vbObjectError + 4006, , "Batch dates require yyyy-mm-dd or null."
                values(key) = DateSerial(CLng(Left$(values(key), 4)), CLng(Mid$(values(key), 6, 2)), CLng(Right$(values(key), 2)))
            End If
        End If
    Next
    For Each key In Array("Title", "MemberName", "DraftNotes")
        If IsNull(values(key)) Then Err.Raise vbObjectError + 4009, , "Review and explicitly fill the null draft field before applying a batch: " & key
    Next
    stage = "Normal draft commit"
    CommitSavedFields = CommitDraft(reportID, values("FormKind"), Nz(values("Title"), ""), Nz(values("MemberName"), ""), values("PeriodStart"), values("PeriodEnd"), Nz(values("DraftNotes"), ""), reason, values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    Err.Raise failure, "CommitSavedFields", message
End Function
Public Function SameFieldValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    SameFieldValue = (IsNull(a) = IsNull(b) And StrComp(Nz(a, ""), Nz(b, ""), vbBinaryCompare) = 0)
End Function
Public Function CurrentRevision(ByVal reportID As Long) As Long
    CurrentRevision = Nz(MajorScalar("SELECT CurrentRevisionID FROM Report WHERE ReportID=" & reportID), 0)
    If CurrentRevision = 0 Then Err.Raise vbObjectError + 4002, , "Select a saved report."
End Function
Public Function NumericTraitAverage(ByVal revisionID As Long) As Variant
    Dim key As Variant, value As Variant, total As Double, count As Long
    For Each key In TraitKeys()
        If Not MajorRevisionHasField(revisionID, CStr(key)) Then NumericTraitAverage = Null: Exit Function
        value = MajorRevisionValue(revisionID, CStr(key))
        If Not IsNull(value) Then
            If value <> "" And value <> "NOB" Then total = total + Val(value): count = count + 1
        End If
    Next
    If count = 0 Then NumericTraitAverage = Null Else NumericTraitAverage = total / count
End Function
Public Function MajorRevisionHasField(ByVal revisionID As Long, ByVal key As String) As Boolean
    MajorRevisionHasField = Not IsNull(MajorScalar("SELECT RevisionValueID FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(key)))
End Function
Public Function MajorRevisionValue(ByVal revisionID As Long, ByVal key As String) As Variant
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(key), dbOpenSnapshot)
    If r.EOF Then
        r.Close
        If key = "CommandEmployment" And RevisionSnapshotSchema(revisionID) < 24 Then MajorRevisionValue = Null: Exit Function
        Err.Raise vbObjectError + 4007, , "Snapshot field missing."
    End If
    If r!IsNullValue Then MajorRevisionValue = Null Else MajorRevisionValue = r!ValueText
    r.Close
End Function
