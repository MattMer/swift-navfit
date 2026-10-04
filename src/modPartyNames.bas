Attribute VB_Name = "modPartyNames"
Option Compare Binary
Option Explicit
Public Function PartyNamesKeys() As Variant
    PartyNamesKeys = Array("MemberFirstName", "MemberMiddleInitial", "MemberLastName", "MemberSuffix", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial", "CounselorLastName", "CounselorFirstInitial", "CounselorMiddleInitial")
End Function
Public Function PartyNamesLabel(ByVal key As String) As String
    Select Case key
        Case "MemberFirstName": PartyNamesLabel = "Member first name"
        Case "MemberMiddleInitial": PartyNamesLabel = "Member middle initial(s)"
        Case "MemberLastName": PartyNamesLabel = "Member last name"
        Case "MemberSuffix": PartyNamesLabel = "Member suffix"
        Case "ReportingSeniorLastName": PartyNamesLabel = "Reporting senior last name"
        Case "ReportingSeniorFirstInitial": PartyNamesLabel = "Reporting senior first initial"
        Case "ReportingSeniorMiddleInitial": PartyNamesLabel = "Reporting senior middle initial(s)"
        Case "CounselorLastName": PartyNamesLabel = "Counselor last name"
        Case "CounselorFirstInitial": PartyNamesLabel = "Counselor first initial"
        Case "CounselorMiddleInitial": PartyNamesLabel = "Counselor middle initial(s)"
    End Select
End Function
Public Function ValidatePartyNames(ByVal values As Variant) As Variant
    Dim i As Long, keys As Variant, limits As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 3100, , "Structured names require ten text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 9 Then Err.Raise vbObjectError + 3100, , "Structured names require ten text or null entries."
    keys = PartyNamesKeys(): limits = Array(27, 3, 27, 10, 18, 1, 3, 20, 1, 3)
    For i = 0 To 9
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 3101, , PartyNamesLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 3102, , PartyNamesLabel(keys(i)) & " exceeds " & limits(i) & " characters; no truncation is performed."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 3103, , PartyNamesLabel(keys(i)) & " must be single-line without tabs."

        End If
    Next
    ValidatePartyNames = values
End Function
Public Function PartyNamesFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = PartyNamesKeys()
    For i = 0 To 9: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    PartyNamesFromEditor = values
End Function
Public Function PartyNamesFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = PartyNamesKeys()
    For i = 0 To 9: values(i) = r.Fields(keys(i)).Value: Next
    PartyNamesFromRecordset = values
End Function
Public Function PartyNamesFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = PartyNamesKeys()
    For i = 0 To 9
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 19 Then r.Close: Err.Raise vbObjectError + 3104, , "Structured names snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 3105, , "Structured names snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    PartyNamesFromRevision = ValidatePartyNames(values)
End Function
Public Function OpenPartyNames() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 3106, , "Open a draft first."
    DoCmd.OpenForm "frmPartyNames"
    For Each key In PartyNamesKeys()
        Forms!frmPartyNames.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmPartyNames!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmPartyNames.SetFocus: OpenPartyNames = True
End Function
Public Function ApplyPartyNames() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 3107, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In PartyNamesKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmPartyNames.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmPartyNames", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyPartyNames = True
End Function
Public Function CancelPartyNames() As Boolean
    DoCmd.Close acForm, "frmPartyNames", acSaveNo
    Forms!frmEditor.SetFocus: CancelPartyNames = True
End Function
