Attribute VB_Name = "modMemberDetails"
Option Compare Binary
Option Explicit
Public Function MemberDetailsKeys() As Variant
    MemberDetailsKeys = Array("MemberSSN", "MemberDoDID", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "DateReported")
End Function
Public Function MemberDetailsLabel(ByVal key As String) As String
    Select Case key
        Case "MemberSSN": MemberDetailsLabel = "Member SSN"
        Case "MemberDoDID": MemberDetailsLabel = "Member DoD ID"
        Case "ComponentActive": MemberDetailsLabel = "Active component"
        Case "ComponentTAR": MemberDetailsLabel = "TAR component"
        Case "ComponentInactive": MemberDetailsLabel = "Inactive component"
        Case "ComponentATADSW": MemberDetailsLabel = "AT/ADSW component"
        Case "DateReported": MemberDetailsLabel = "Date reported"
    End Select
End Function
Private Function DigitsOnly(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    DigitsOnly = True
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then DigitsOnly = False: Exit Function
    Next
End Function
Public Function IsValidIsoDate(ByVal text As String) As Boolean
    Dim d As Date
    On Error GoTo Invalid
    If Len(text) <> 10 Or Mid$(text, 5, 1) <> "-" Or Mid$(text, 8, 1) <> "-" Then Exit Function
    If Not DigitsOnly(Left$(text, 4) & Mid$(text, 6, 2) & Right$(text, 2)) Then Exit Function
    d = DateSerial(CLng(Left$(text, 4)), CLng(Mid$(text, 6, 2)), CLng(Right$(text, 2)))
    IsValidIsoDate = (Format$(d, "yyyy-mm-dd") = text)
Invalid:
End Function
Public Function ValidateMemberDetails(ByVal values As Variant) As Variant
    Dim i As Long, choice As Variant, allowed As Boolean, limit As Long
    If Not IsArray(values) Then Err.Raise vbObjectError + 2600, , "Member details require seven text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 6 Then Err.Raise vbObjectError + 2600, , "Member details require seven text or null entries."
    For i = 0 To 6
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2601, , "Member details must be text or null."
            If i < 2 Then
                If i = 0 Then limit = 9 Else limit = 10
                If Len(values(i)) > limit Or Not DigitsOnly(CStr(values(i))) Then Err.Raise vbObjectError + 2602, , "Identifiers accept digits only, up to their storage capacity. Leading zeros remain intact; no formatting or truncation is performed."
            ElseIf i < 6 Then
                allowed = False
                For Each choice In Array("", "Yes", "No")
                    If values(i) = choice Then allowed = True
                Next
                If Not allowed Then Err.Raise vbObjectError + 2603, , "Choose Yes, No or blank for each component selection."
            ElseIf Len(values(i)) > 0 Then
                If Not IsValidIsoDate(CStr(values(i))) Then Err.Raise vbObjectError + 2604, , "Date reported must be a real date in yyyy-mm-dd format, or blank."
            End If
        End If
    Next
    ValidateMemberDetails = values
End Function
Public Function MemberDetailsFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null): keys = MemberDetailsKeys()
    For i = 0 To 6: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    MemberDetailsFromEditor = values
End Function
Public Function MemberDetailsFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null): keys = MemberDetailsKeys()
    For i = 0 To 6: values(i) = r.Fields(keys(i)).Value: Next
    MemberDetailsFromRecordset = values
End Function
Public Function MemberDetailsFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null): keys = MemberDetailsKeys()
    For i = 0 To 6
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 15 Then r.Close: Err.Raise vbObjectError + 2604, , "Member details snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2605, , "Member details snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    MemberDetailsFromRevision = ValidateMemberDetails(values)
End Function
Public Function OpenMemberDetails() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2606, , "Open a draft first."
    DoCmd.OpenForm "frmMemberDetails"
    For Each key In MemberDetailsKeys()
        Forms!frmMemberDetails.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmMemberDetails!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmMemberDetails.SetFocus: OpenMemberDetails = True
End Function
Public Function ApplyMemberDetails() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2607, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In MemberDetailsKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmMemberDetails.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmMemberDetails", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyMemberDetails = True
End Function
Public Function CancelMemberDetails() As Boolean
    DoCmd.Close acForm, "frmMemberDetails", acSaveNo
    Forms!frmEditor.SetFocus: CancelMemberDetails = True
End Function
