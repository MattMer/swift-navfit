Attribute VB_Name = "modPromotionStatus"
Option Compare Binary
Option Explicit
Public Function ValidatePromotionStatus(ByVal value As Variant) As Variant
    Dim choice As Variant, allowed As Boolean
    If Not IsNull(value) Then
        If VarType(value) <> vbString Then Err.Raise vbObjectError + 2300, "modPromotionStatus.ValidatePromotionStatus", "Promotion status must be a listed text selection or blank."
        For Each choice In Array("", "REGULAR", "FROCKED", "SELECTED", "SPOT")
            If value = choice Then allowed = True
        Next
        If Not allowed Then Err.Raise vbObjectError + 2301, "modPromotionStatus.ValidatePromotionStatus", "Choose REGULAR, FROCKED, SELECTED or SPOT, or clear promotion status."
    End If
    ValidatePromotionStatus = value
End Function
' Block 8 as read from a NAVFIT98A file. NAVFIT98A stores whatever was typed, so case and
' surrounding spaces are ignored; any other value comes in blank rather than stopping the
' import (block 8 then shows as incomplete, and the import receipt keeps the original).
Public Function ImportedPromotionStatus(ByVal value As Variant) As Variant
    Dim s As String
    ImportedPromotionStatus = Null
    If IsNull(value) Then Exit Function
    'An empty block 8 stays empty, so it goes back to NAVFIT98A exactly as it came.
    If CStr(value) = "" Then ImportedPromotionStatus = "": Exit Function
    s = UCase$(Trim$(CStr(value)))
    Select Case s
        Case "REGULAR", "FROCKED", "SELECTED", "SPOT": ImportedPromotionStatus = s
    End Select
End Function
Public Function PromotionStatusFromRevision(ByVal revisionID As Long) As Variant
    Dim r As DAO.Recordset, value As Variant
    value = Null
    Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='PromotionStatus'", dbOpenSnapshot)
    If r.EOF Then
        If RevisionSnapshotSchema(revisionID) >= 12 And Not RevisionIsSparse(revisionID) Then r.Close: Err.Raise vbObjectError + 2302, "modPromotionStatus.PromotionStatusFromRevision", "Promotion status snapshot is incomplete."
    ElseIf Not r!IsNullValue Then
        If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2303, "modPromotionStatus.PromotionStatusFromRevision", "Promotion status snapshot has an inconsistent null flag."
        value = r!ValueText
    End If
    r.Close
    PromotionStatusFromRevision = ValidatePromotionStatus(value)
End Function
