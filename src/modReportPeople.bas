Attribute VB_Name = "modReportPeople"
Option Compare Binary
Option Explicit
Public Function ReportPeopleKeys() As Variant
    ReportPeopleKeys = Array("ReportingSeniorSSN", "ReportingSeniorDoDID", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "RaterName", "RaterDate", "SeniorRaterName", "SeniorRaterDate")
End Function
Public Function ReportPeopleLabel(ByVal key As String) As String
    Select Case key
        Case "ReportingSeniorSSN": ReportPeopleLabel = "Reporting senior SSN"
        Case "ReportingSeniorDoDID": ReportPeopleLabel = "Reporting senior DoD ID"
        Case "ReportingSeniorAddress1": ReportPeopleLabel = "Address line 1"
        Case "ReportingSeniorAddress2": ReportPeopleLabel = "Address line 2"
        Case "ReportingSeniorCity": ReportPeopleLabel = "City"
        Case "ReportingSeniorState": ReportPeopleLabel = "State / region"
        Case "ReportingSeniorZip": ReportPeopleLabel = "ZIP / postal code"
        Case "ReportingSeniorPhone": ReportPeopleLabel = "Phone"
        Case "ReportingSeniorDSN": ReportPeopleLabel = "DSN"
        Case "RaterName": ReportPeopleLabel = "Rater name"
        Case "RaterDate": ReportPeopleLabel = "Rater date"
        Case "SeniorRaterName": ReportPeopleLabel = "Senior rater name"
        Case "SeniorRaterDate": ReportPeopleLabel = "Senior rater date"
    End Select
End Function
Private Function ReportingIdentifierDigits(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    ReportingIdentifierDigits = True
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch < "0" Or ch > "9" Then ReportingIdentifierDigits = False: Exit Function
    Next
End Function
Public Function ValidateReportPeople(ByVal values As Variant) As Variant
    Dim i As Long, keys As Variant, limits As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 2900, , "Reporting contact and rater details require thirteen text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 12 Then Err.Raise vbObjectError + 2900, , "Reporting contact and rater details require thirteen text or null entries."
    keys = ReportPeopleKeys(): limits = CatalogueFieldLimits(keys)
    For i = 0 To 12
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2901, , ReportPeopleLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 2902, , ReportPeopleLabel(keys(i)) & " exceeds " & limits(i) & " characters; no truncation is performed."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 2903, , ReportPeopleLabel(keys(i)) & " must be single-line without tabs."
            If i < 2 Then
                If Not ReportingIdentifierDigits(CStr(values(i))) Then Err.Raise vbObjectError + 2904, , ReportPeopleLabel(keys(i)) & " accepts digits only. Leading zeros are retained."
            ElseIf (i = 10 Or i = 12) And Len(values(i)) > 0 Then
                If Not IsValidIsoDate(CStr(values(i))) Then Err.Raise vbObjectError + 2905, , ReportPeopleLabel(keys(i)) & " must be a real yyyy-mm-dd date or blank."
            End If
        End If
    Next
    ValidateReportPeople = values
End Function
Public Function ReportPeopleFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportPeopleKeys()
    For i = 0 To 12: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    ReportPeopleFromEditor = values
End Function
Public Function ReportPeopleFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportPeopleKeys()
    For i = 0 To 12: values(i) = r.Fields(keys(i)).Value: Next
    ReportPeopleFromRecordset = values
End Function
Public Function ReportPeopleFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportPeopleKeys()
    For i = 0 To 12
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 18 Then r.Close: Err.Raise vbObjectError + 2904, , "Reporting contact and rater details snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2905, , "Reporting contact and rater details snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    ReportPeopleFromRevision = ValidateReportPeople(values)
End Function
Public Function OpenReportPeople() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2906, , "Open a draft first."
    DoCmd.OpenForm "frmReportPeople"
    For Each key In ReportPeopleKeys()
        Forms!frmReportPeople.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmReportPeople!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmReportPeople.SetFocus: OpenReportPeople = True
End Function
Public Function ApplyReportPeople() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2907, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In ReportPeopleKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmReportPeople.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmReportPeople", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyReportPeople = True
End Function
Public Function CancelReportPeople() As Boolean
    DoCmd.Close acForm, "frmReportPeople", acSaveNo
    Forms!frmEditor.SetFocus: CancelReportPeople = True
End Function
