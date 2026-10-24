Attribute VB_Name = "modAssessment"
Option Compare Binary
Option Explicit
Public Function AssessmentKeys() As Variant
    AssessmentKeys = Array("ObservationStatus", "PromotionRecommendation")
End Function
Public Function AssessmentLabel(ByVal key As String) As String
    Select Case key
        Case "ObservationStatus": AssessmentLabel = "Whole-report observation status"
        Case "PromotionRecommendation": AssessmentLabel = "Promotion recommendation"
    End Select
End Function
Public Function ValidateAssessment(ByVal values As Variant) As Variant
    Dim i As Long, choice As Variant, choices As Variant, allowed As Boolean
    If Not IsArray(values) Then Err.Raise vbObjectError + 2200, "modAssessment.ValidateAssessment", "Assessment must contain two text or null selections."
    If LBound(values) <> 0 Or UBound(values) <> 1 Then Err.Raise vbObjectError + 2200, "modAssessment.ValidateAssessment", "Assessment must contain two text or null selections."
    For i = 0 To 1
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2201, "modAssessment.ValidateAssessment", "Assessment selections must be text or null."
            If i = 0 Then
                choices = Array("", "Observed", "NOB")
            Else
                choices = Array("", "NOB", "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
            End If
            allowed = False
            For Each choice In choices
                If values(i) = choice Then allowed = True
            Next
            If Not allowed Then Err.Raise vbObjectError + 2202, "modAssessment.ValidateAssessment", "Choose a listed assessment selection or clear it. No assessment is inferred from trait marks."
        End If
    Next
    ValidateAssessment = values
End Function
Public Function AssessmentFromEditor() As Variant
    AssessmentFromEditor = Array(Forms!frmEditor!cboObservationStatus.Value, Forms!frmEditor!cboPromotionRecommendation.Value)
End Function
Public Function AssessmentFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long, schema As Long
    values = Array(Null, Null): keys = AssessmentKeys(): schema = RevisionSnapshotSchema(revisionID)
    For i = 0 To 1
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If schema >= 11 Then r.Close: Err.Raise vbObjectError + 2203, "modAssessment.AssessmentFromRevision", "Assessment snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2204, "modAssessment.AssessmentFromRevision", "Assessment snapshot has an inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    AssessmentFromRevision = ValidateAssessment(values)
End Function
