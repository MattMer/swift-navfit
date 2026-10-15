Attribute VB_Name = "modUXGlobalDefaults"
Option Explicit
Public Function UXDefaultScopeWhere(ByVal groupID As Long) As String
    If groupID = 0 Then UXDefaultScopeWhere = "ScopeKind='Command' AND ScopeID=1" Else UXDefaultScopeWhere = "ScopeKind='Group' AND ScopeID=" & groupID
End Function
Public Function UXBulkContextStamp(ByVal groupID As Long) As String
    If groupID = 0 Then UXBulkContextStamp = "Global:" & UXDefaultsStamp(0) Else UXBulkContextStamp = GroupStamp(groupID)
End Function
Public Function UXOpenGlobalDefaults() As Boolean
    UXOpenGlobalDefaults = UXStartGroupBulk(True, True)
End Function
Public Function UXSeedGlobalDefaults() As Boolean
    Dim key As Variant, name As String, value As Variant
    For Each key In UXSharedKeys()
        name = UXFieldControl(CStr(key))
        If Len(name) > 0 Then
            value = UXGroupDefaultValue(0, CStr(key))
            If Not IsNull(value) Then
                If Len(Nz(Forms!frmEditor.Controls(name), "")) = 0 Then Forms!frmEditor.Controls(name) = value
            End If
        End If
    Next
    UXSeedGlobalDefaults = True
End Function
Public Function UXEffectiveDefault(ByVal groupID As Long, ByVal key As String) As Variant
    If groupID > 0 And DCount("*", "FieldDefault", UXDefaultScopeWhere(groupID) & " AND FieldKey=" & QuoteSQL(key)) = 0 Then UXEffectiveDefault = UXGroupDefaultValue(0, key) Else UXEffectiveDefault = UXGroupDefaultValue(groupID, key)
End Function
Public Sub UXRemoveDefault(ByVal groupID As Long, ByVal key As String)
    Dim op As Long, old As Variant
    old = UXGroupDefaultValue(groupID, key)
    op = BeginOperation("Reset default override", UXDefaultScopeWhere(groupID) & "; field " & key & "; before=" & Nz(old, "(null or absent)"))
    CurrentDb.Execute "DELETE FROM FieldDefault WHERE " & UXDefaultScopeWhere(groupID) & " AND FieldKey=" & QuoteSQL(key), dbFailOnError
    FinishOperation op, "Completed", "Explicit override removed; inherit parent starting value. Saved reports unchanged."
End Sub
