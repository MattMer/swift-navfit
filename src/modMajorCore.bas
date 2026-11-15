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
    If Len(Trim$(csv)) = 0 Then Err.Raise vbObjectError + 4001, "modMajorCore.MajorIDs", "Select at least one report."
    For Each item In Split(csv, ",")
        value = Trim$(CStr(item))
        If Len(value) = 0 Or Len(value) > 9 Then Err.Raise vbObjectError + 4001, "modMajorCore.MajorIDs", "Invalid report selection."
        For i = 1 To Len(value)
            If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Err.Raise vbObjectError + 4001, "modMajorCore.MajorIDs", "Invalid report selection."
        Next
        n = CLng(value)
        If n < 1 Then Err.Raise vbObjectError + 4001, "modMajorCore.MajorIDs", "Invalid report selection."
        If result.Exists(CStr(n)) Then Err.Raise vbObjectError + 4001, "modMajorCore.MajorIDs", "Duplicate report selection."
        result.Add CStr(n), n
    Next
    Set MajorIDs = result
End Function
Public Function SavedFields(ByVal reportID As Long) As Object
    Dim r As DAO.Recordset, values As Object, key As Variant, db As DAO.Database
    AssertReportCache reportID
    Set db = CurrentDb
    Set values = CreateObject("Scripting.Dictionary")
    Dim head As Long, value As Variant, block28Recorded As Boolean
    head = ReportHeadRevision(reportID)
    For Each key In DraftFieldKeys(): values.Add CStr(key), Null: Next
    Set r = db.OpenRecordset("SELECT FieldKey,ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & head, dbOpenSnapshot)
    Do While Not r.EOF
        key = CStr(r!FieldKey): value = Null
        If Not r!IsNullValue Then value = r!ValueText.Value
        If (key = "PeriodStart" Or key = "PeriodEnd") And Not IsNull(value) Then value = DateSerial(CLng(Left$(value, 4)), CLng(Mid$(value, 6, 2)), CLng(Right$(value, 2)))
        If values.Exists(CStr(key)) Then values(CStr(key)) = value
        If key = "CommandEmployment" Then block28Recorded = True
        r.MoveNext
    Loop
    r.Close
    ' Revisions saved before Block 28 existed carry its native imported value as a
    ' read-only projection, so exports and later saves do not clear it.
    If Not block28Recorded And Not RevisionIsSparse(head) Then values("CommandEmployment") = StoredBlock28(reportID)
    Set SavedFields = values
End Function
Public Function FieldsSubset(ByVal values As Object, ByVal keys As Variant) As Variant
    Dim result() As Variant, i As Long
    ReDim result(0 To UBound(keys))
    For i = 0 To UBound(keys): result(i) = values(CStr(keys(i))): Next
    FieldsSubset = result
End Function
Public Function CommitSavedFields(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal changes As Object, ByVal reason As String, Optional ByVal ownTransaction As Boolean = True, Optional ByVal syncDirectory As Boolean = True) As Long
    'False is reserved for callers that already own the entire guarded batch.
    Dim db As DAO.Database, r As DAO.Recordset, values As Object, key As Variant, w As DAO.Workspace, failure As Long, message As String, stage As String
    On Error GoTo Failed
    stage = "Begin guarded transaction"
    Set w = DBEngine.Workspaces(0)
    If ownTransaction Then w.BeginTrans
    stage = "Revision guard": Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT CurrentRevisionID,IsArchived,WorkflowState FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 4002, "modMajorCore.CommitSavedFields", "Report does not exist."
    If r!IsArchived Or r!WorkflowState <> "Draft" Then r.Close: Err.Raise vbObjectError + 4003, "modMajorCore.CommitSavedFields", "Archived or locked report cannot be changed."
    If expectedRevision = 0 Or Nz(r!CurrentRevisionID, 0) <> expectedRevision Then r.Close: Err.Raise vbObjectError + 4004, "modMajorCore.CommitSavedFields", "Report changed since preview."
    r.Close: stage = "Read saved fields": Set values = SavedFields(reportID)
    For Each key In changes.Keys
        If Not values.Exists(key) Or key = "FormKind" Then Err.Raise vbObjectError + 4005, "modMajorCore.CommitSavedFields", "Unknown or immutable draft field."
        If (key = "Title" Or key = "MemberName" Or key = "DraftNotes") And IsNull(changes(key)) Then Err.Raise vbObjectError + 4008, "modMajorCore.CommitSavedFields", "This draft field uses empty text. Clear the null selection and use an empty value."
        values(key) = changes(key)
        If key = "PeriodStart" Or key = "PeriodEnd" Then
            If Not IsNull(values(key)) And VarType(values(key)) <> vbDate Then
                If Not IsValidIsoDate(CStr(values(key))) Then Err.Raise vbObjectError + 4006, "modMajorCore.CommitSavedFields", "Batch dates require yyyy-mm-dd or null."
                values(key) = DateSerial(CLng(Left$(values(key), 4)), CLng(Mid$(values(key), 6, 2)), CLng(Right$(values(key), 2)))
            End If
        End If
    Next
    For Each key In Array("Title", "MemberName", "DraftNotes")
        If IsNull(values(key)) Then Err.Raise vbObjectError + 4009, "modMajorCore.CommitSavedFields", "Review and explicitly fill the null draft field before applying a batch: " & key
    Next
    stage = "Normal draft commit"
    CommitSavedFields = CommitDraft(reportID, values("FormKind"), Nz(values("Title"), ""), Nz(values("MemberName"), ""), values("PeriodStart"), values("PeriodEnd"), Nz(values("DraftNotes"), ""), reason, values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"), values, syncDirectory:=syncDirectory)
    If ownTransaction Then w.CommitTrans
    Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modMajorCore", "CommitSavedFields", stage, Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = stage & ": " & CStr(swiftPrimary("Description"))
    If ownTransaction Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function SameFieldValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    SameFieldValue = (IsNull(a) = IsNull(b) And StrComp(Nz(a, ""), Nz(b, ""), vbBinaryCompare) = 0)
End Function
Public Function CurrentRevision(ByVal reportID As Long) As Long
    CurrentRevision = ReportHeadRevision(reportID)
    If CurrentRevision = 0 Then Err.Raise vbObjectError + 4002, "modMajorCore.CurrentRevision", "Select a saved report."
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
' A revision saved from schema 50 is sparse (ValueEncoding 1): only fields with a value
' have a row, and a missing field is empty. Older revisions are full copies.
Public Function RevisionIsSparse(ByVal revisionID As Long) As Boolean
    RevisionIsSparse = (Nz(MajorScalar("SELECT ValueEncoding FROM ReportRevision WHERE RevisionID=" & revisionID), 0) = 1)
End Function
' A sparse revision holds every catalogue field that existed when it was saved, so a
' missing one is empty; a field added to the catalogue later was not recorded.
Public Function SparseHoldsField(ByVal revisionID As Long, ByVal key As String) As Boolean
    Dim r As DAO.Recordset, version As Long
    Set r = CurrentDb.OpenRecordset("SELECT ValueEncoding,FieldDefinitionVersion FROM ReportRevision WHERE RevisionID=" & revisionID, dbOpenSnapshot)
    If Not r.EOF Then
        If Nz(r!ValueEncoding, 0) = 1 Then
            version = CatalogueFieldVersion(key)
            SparseHoldsField = (version >= 1 And version <= Nz(r!FieldDefinitionVersion, 0))
        End If
    End If
    r.Close
End Function
' The one way to read a single saved field. Returns False when the snapshot does not
' hold the field (a full copy without its row); value is Null for an empty field.
Public Function RevisionStoredValue(ByVal revisionID As Long, ByVal key As String, ByRef value As Variant) As Boolean
    Dim r As DAO.Recordset
    value = Null
    Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(key), dbOpenSnapshot)
    If r.EOF Then
        RevisionStoredValue = SparseHoldsField(revisionID, key)
    Else
        If Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 4008, "modMajorCore.RevisionStoredValue", "Snapshot has an inconsistent null flag."
            value = r!ValueText.Value
        End If
        RevisionStoredValue = True
    End If
    r.Close
End Function
' A sparse snapshot read as rows lists only fields with a value; add the empty ones
' that existed at the revision's catalogue version.
Public Sub CompleteSparseSnapshot(ByVal values As Object, ByVal fieldVersion As Long)
    Dim key As Variant, version As Long
    For Each key In DraftFieldKeys()
        version = CatalogueFieldVersion(CStr(key))
        If Not values.Exists(CStr(key)) And version >= 1 And version <= fieldVersion Then values.Add CStr(key), Null
    Next
End Sub
Public Function MajorRevisionHasField(ByVal revisionID As Long, ByVal key As String) As Boolean
    Dim value As Variant
    MajorRevisionHasField = RevisionStoredValue(revisionID, key, value)
End Function
Public Function MajorRevisionValue(ByVal revisionID As Long, ByVal key As String) As Variant
    Dim value As Variant
    If RevisionStoredValue(revisionID, key, value) Then MajorRevisionValue = value: Exit Function
    If key = "CommandEmployment" And RevisionSnapshotSchema(revisionID) < 24 Then MajorRevisionValue = Null: Exit Function
    Err.Raise vbObjectError + 4007, "modMajorCore.MajorRevisionValue", "Snapshot field missing."
End Function
