Attribute VB_Name = "modServiceDetails"
Option Compare Binary
Option Explicit
Public Function IsCounselingStatus(ByVal value As String) As Boolean
    IsCounselingStatus = (value = "NOT REQ" Or value = "NOT PERF")
End Function
Public Function ServiceDetailsKeys() As Variant
    ServiceDetailsKeys = Array("CounselingDate", "CounselorName", "ReadinessSummary", "ReadinessDetail", "ReadinessDate", "BilletSubcategory")
End Function
Public Function ServiceDetailsLabel(ByVal key As String) As String
    Select Case key
        Case "CounselingDate": ServiceDetailsLabel = "Counseling date"
        Case "CounselorName": ServiceDetailsLabel = "Counselor name"
        Case "ReadinessSummary": ServiceDetailsLabel = "Physical-readiness summary"
        Case "ReadinessDetail": ServiceDetailsLabel = "Physical-readiness detail"
        Case "ReadinessDate": ServiceDetailsLabel = "Physical-readiness date"
        Case "BilletSubcategory": ServiceDetailsLabel = "Billet subcategory"
    End Select
End Function
Public Function ValidateServiceDetails(ByVal values As Variant) As Variant
    Dim i As Long, keys As Variant, limits As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 2700, , "Service details require six text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 5 Then Err.Raise vbObjectError + 2700, , "Service details require six text or null entries."
    keys = ServiceDetailsKeys(): limits = CatalogueFieldLimits(keys)
    For i = 0 To 5
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2701, , ServiceDetailsLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 2702, , ServiceDetailsLabel(keys(i)) & " exceeds " & limits(i) & " characters; no truncation is performed."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 2703, , ServiceDetailsLabel(keys(i)) & " must be single-line without tabs."
            If (i = 0 Or i = 4) And Len(values(i)) > 0 Then
                If Not IsValidIsoDate(CStr(values(i))) And Not (i = 0 And IsCounselingStatus(CStr(values(i)))) Then Err.Raise vbObjectError + 2704, , ServiceDetailsLabel(keys(i)) & " must be a real ISO date, or blank (counseling also accepts NOT REQ / NOT PERF)."
            End If
        End If
    Next
    ValidateServiceDetails = values
End Function
Public Function ServiceDetailsFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = ServiceDetailsKeys()
    For i = 0 To 5: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    ServiceDetailsFromEditor = values
End Function
Public Function ServiceDetailsFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = ServiceDetailsKeys()
    For i = 0 To 5: values(i) = r.Fields(keys(i)).Value: Next
    ServiceDetailsFromRecordset = values
End Function
Public Function ServiceDetailsFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = ServiceDetailsKeys()
    For i = 0 To 5
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 16 Then r.Close: Err.Raise vbObjectError + 2704, , "Service details snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2705, , "Service details snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    ServiceDetailsFromRevision = ValidateServiceDetails(values)
End Function
Public Function OpenServiceDetails() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2706, , "Open a draft first."
    DoCmd.OpenForm "frmServiceDetails"
    For Each key In ServiceDetailsKeys()
        Forms!frmServiceDetails.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmServiceDetails!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmServiceDetails.SetFocus: OpenServiceDetails = True
End Function
Public Function ApplyServiceDetails() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2707, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In ServiceDetailsKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmServiceDetails.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmServiceDetails", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyServiceDetails = True
End Function
Public Function CancelServiceDetails() As Boolean
    DoCmd.Close acForm, "frmServiceDetails", acSaveNo
    Forms!frmEditor.SetFocus: CancelServiceDetails = True
End Function
