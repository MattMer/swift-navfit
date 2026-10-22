Attribute VB_Name = "modReportingSenior"
Option Compare Binary
Option Explicit
Public Function ReportingSeniorKeys() As Variant
    ReportingSeniorKeys = Array("ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC")
End Function
Public Function ReportingSeniorLabel(ByVal key As String) As String
    Select Case key
        Case "ReportingSeniorName": ReportingSeniorLabel = "Reporting senior name"
        Case "ReportingSeniorGrade": ReportingSeniorLabel = "Reporting senior grade"
        Case "ReportingSeniorDesignator": ReportingSeniorLabel = "Reporting senior designator"
        Case "ReportingSeniorTitle": ReportingSeniorLabel = "Reporting senior title"
        Case "ReportingSeniorUIC": ReportingSeniorLabel = "Reporting senior UIC"
    End Select
End Function
Public Function ValidateReportingSenior(ByVal values As Variant) As Variant
    Dim i As Long, limits As Variant, keys As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 2400, , "Reporting senior requires five text or null values."
    If LBound(values) <> 0 Or UBound(values) <> 4 Then Err.Raise vbObjectError + 2400, , "Reporting senior requires five text or null values."
    keys = ReportingSeniorKeys(): limits = CatalogueFieldLimits(keys)
    For i = 0 To 4
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2401, , ReportingSeniorLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 2402, , ReportingSeniorLabel(keys(i)) & " exceeds " & limits(i) & " characters; shorten it deliberately."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 2403, , ReportingSeniorLabel(keys(i)) & " must be single-line text without tabs."
        End If
    Next
    ValidateReportingSenior = values
End Function
Public Function ReportingSeniorFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null): keys = ReportingSeniorKeys()
    For i = 0 To 4: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    ReportingSeniorFromEditor = values
End Function
Public Function ReportingSeniorFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null): keys = ReportingSeniorKeys()
    For i = 0 To 4
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 13 Then r.Close: Err.Raise vbObjectError + 2404, , "Reporting senior snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2405, , "Reporting senior snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    ReportingSeniorFromRevision = ValidateReportingSenior(values)
End Function
Public Function OpenReportingSenior() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2406, , "Open a draft first."
    DoCmd.OpenForm "frmReportingSenior"
    For Each key In ReportingSeniorKeys()
        Forms!frmReportingSenior.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmReportingSenior!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmReportingSenior.SetFocus: OpenReportingSenior = True
End Function
Public Function ApplyReportingSenior() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2407, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In ReportingSeniorKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmReportingSenior.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmReportingSenior", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyReportingSenior = True
End Function
Public Function CancelReportingSenior() As Boolean
    DoCmd.Close acForm, "frmReportingSenior", acSaveNo
    Forms!frmEditor.SetFocus: CancelReportingSenior = True
End Function
