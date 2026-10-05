Attribute VB_Name = "modPromotionStatus"
Option Compare Binary
Option Explicit
Public Function ValidatePromotionStatus(ByVal value As Variant) As Variant
    Dim choice As Variant, allowed As Boolean
    If Not IsNull(value) Then
        If VarType(value) <> vbString Then Err.Raise vbObjectError + 2300, , "Promotion status must be a listed text selection or blank."
        For Each choice In Array("", "REGULAR", "FROCKED", "SELECTED", "SPOT")
            If value = choice Then allowed = True
        Next
        If Not allowed Then Err.Raise vbObjectError + 2301, , "Choose REGULAR, FROCKED, SELECTED or SPOT, or clear promotion status."
    End If
    ValidatePromotionStatus = value
End Function
Public Function PromotionStatusFromRevision(ByVal revisionID As Long) As Variant
    Dim r As DAO.Recordset, value As Variant
    value = Null
    Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='PromotionStatus'", dbOpenSnapshot)
    If r.EOF Then
        If RevisionSnapshotSchema(revisionID) >= 12 Then r.Close: Err.Raise vbObjectError + 2302, , "Promotion status snapshot is incomplete."
    ElseIf Not r!IsNullValue Then
        If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2303, , "Promotion status snapshot has an inconsistent null flag."
        value = r!ValueText
    End If
    r.Close
    PromotionStatusFromRevision = ValidatePromotionStatus(value)
End Function
