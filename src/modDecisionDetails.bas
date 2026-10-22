Attribute VB_Name = "modDecisionDetails"
Option Compare Binary
Option Explicit
Public Function DecisionDetailsKeys() As Variant
    DecisionDetailsKeys = Array("AssignmentRecommendationA", "AssignmentRecommendationB", "RetentionYes", "RetentionNo", "StatementYes", "StatementNo")
End Function
Public Function DecisionDetailsLabel(ByVal key As String) As String
    Select Case key
        Case "AssignmentRecommendationA": DecisionDetailsLabel = "Assignment recommendation A"
        Case "AssignmentRecommendationB": DecisionDetailsLabel = "Assignment recommendation B"
        Case "RetentionYes": DecisionDetailsLabel = "Recommended for retention: Yes"
        Case "RetentionNo": DecisionDetailsLabel = "Recommended for retention: No"
        Case "StatementYes": DecisionDetailsLabel = "Member statement: Yes"
        Case "StatementNo": DecisionDetailsLabel = "Member statement: No"
    End Select
End Function
Public Function ValidateDecisionDetails(ByVal values As Variant) As Variant
    Dim i As Long, keys As Variant, limits As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 2800, , "Recommendation and statement details require six text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 5 Then Err.Raise vbObjectError + 2800, , "Recommendation and statement details require six text or null entries."
    keys = DecisionDetailsKeys(): limits = CatalogueFieldLimits(keys)
    For i = 0 To 5
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2801, , DecisionDetailsLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 2802, , DecisionDetailsLabel(keys(i)) & " exceeds " & limits(i) & " characters; no truncation is performed."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 2803, , DecisionDetailsLabel(keys(i)) & " must be single-line without tabs."
            If i >= 2 And Len(values(i)) > 0 Then
                If values(i) <> "Yes" And values(i) <> "No" Then Err.Raise vbObjectError + 2804, , DecisionDetailsLabel(keys(i)) & " must be Yes, No or blank."
            End If
        End If
    Next
    ValidateDecisionDetails = values
End Function
Public Function DecisionDetailsFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = DecisionDetailsKeys()
    For i = 0 To 5: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    DecisionDetailsFromEditor = values
End Function
Public Function DecisionDetailsFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = DecisionDetailsKeys()
    For i = 0 To 5: values(i) = r.Fields(keys(i)).Value: Next
    DecisionDetailsFromRecordset = values
End Function
Public Function DecisionDetailsFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null): keys = DecisionDetailsKeys()
    For i = 0 To 5
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 17 Then r.Close: Err.Raise vbObjectError + 2804, , "Recommendation and statement details snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2805, , "Recommendation and statement details snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    DecisionDetailsFromRevision = ValidateDecisionDetails(values)
End Function
Public Function OpenDecisionDetails() As Boolean
    Dim key As Variant
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2806, , "Open a draft first."
    DoCmd.OpenForm "frmDecisionDetails"
    For Each key In DecisionDetailsKeys()
        Forms!frmDecisionDetails.Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms!frmDecisionDetails!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms!frmDecisionDetails.SetFocus: OpenDecisionDetails = True
End Function
Public Function ApplyDecisionDetails() As Boolean
    Dim key As Variant
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2807, , "Recover the archived draft before editing."
    'Transfer raw pending values; full validation occurs at the atomic draft save.
    For Each key In DecisionDetailsKeys()
        Forms!frmEditor.Controls("txt" & key) = Forms!frmDecisionDetails.Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, "frmDecisionDetails", acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyDecisionDetails = True
End Function
Public Function CancelDecisionDetails() As Boolean
    DoCmd.Close acForm, "frmDecisionDetails", acSaveNo
    Forms!frmEditor.SetFocus: CancelDecisionDetails = True
End Function
