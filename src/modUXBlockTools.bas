Attribute VB_Name = "modUXBlockTools"
Option Explicit
Private blockSection As Long
Public Function UXBlockRefresh() As Boolean
    Dim f As Access.Form, key As Variant
    Set f = Forms!frmUXBulk!uxBlocks.Form
    For Each key In UXSharedKeys()
        f.Controls("value" & key) = Forms!frmUXBulk.Controls("uxValue" & key).Value
    Next
    UXBlockRefresh = True
End Function
Public Function UXBlockEdit(ByVal key As String) As Boolean
    Forms!frmUXBulk.Controls("uxValue" & key) = UXDateEntry(Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key).Value)
    UXBlockEdit = UXConfirmBlockChange(key)
End Function
Public Function UXBlockClear(ByVal key As String) As Boolean
    Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key) = Null
    UXBlockClear = UXBlockEdit(key)
End Function
Public Function UXBlockLayout(ByVal section As Long) As Boolean
    Dim f As Access.Form, c As Access.Control, key As Variant, kind As String, x As Long, y As Long, usable As Long, size As Long, rowHeight As Long, band As String, prior As String, flag As Boolean, multi As Boolean, tabOrder As Long
    Set f = Forms!frmUXBulk!uxBlocks.Form: blockSection = section
    kind = Nz(Forms!frmUXBulk!uxKind, "EVAL"): usable = Forms!frmUXBulk!uxBlocks.Width - 700
    f!anchor.SetFocus
    For Each c In f.Controls
        If c.Name <> "anchor" Then c.Visible = False: c.Left = 100: c.Top = 100: c.Width = 200
    Next
    x = 150: y = 150
    For Each key In UXLayoutKeys(section, kind)
        If UXHasControl(f, "value" & key) And UXFieldApplies(key, kind) And Not UXFieldLegacy(CStr(key)) Then
            flag = UXFieldFlag(key): multi = UXFieldMultiline(key): band = UXLayoutBand(CStr(key))
            size = UXLayoutWidth(CStr(key), usable, kind)
            If size < 2000 Then size = 2000
            If multi Then size = usable: band = CStr(key)
            If x > 150 And (band <> prior Or x + size > usable + 150) Then y = y + rowHeight: x = 150: rowHeight = 0
            With f.Controls("label" & key)
                .Caption = UXFieldCaption(key, kind): .Left = x: .Top = y: .Width = size: .Height = 300: .Visible = True
            End With
            With f.Controls("value" & key)
                .Left = x: .Top = y + 320: .Width = size: .Height = IIf(multi, 1800, 360): .Visible = True: .TabIndex = tabOrder
                tabOrder = tabOrder + 1
                If flag Then .Top = y: .Width = 280: .Height = 280
            End With
            If flag Then
                f.Controls("label" & key).Left = x + 360: f.Controls("label" & key).Width = size - 360
            End If
            If multi Then
                rowHeight = 2220
            ElseIf rowHeight < IIf(flag, 380, 760) Then
                rowHeight = IIf(flag, 380, 760)
            End If
            x = x + size + 160: prior = band
        End If
    Next
    f.Width = usable + 300: f.Section(acDetail).Height = y + rowHeight + 200
    UXBlockLayout = True
End Function
Public Function UXBlockResize() As Boolean
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then UXBlockResize = UXBlockLayout(blockSection)
End Function
Public Function UXDefaultsStamp(ByVal groupID As Long) As String
    Dim r As DAO.Recordset, db As DAO.Database, result As String
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM FieldDefault WHERE " & UXDefaultScopeWhere(groupID) & " ORDER BY FieldKey", dbOpenSnapshot)
    Do While Not r.EOF
        result = result & r!FieldKey & ":" & r!VersionCounter & ":" & Nz(r!ValueText, "") & ":" & r!IsNullValue & "|"
        r.MoveNext
    Loop
    r.Close
    If groupID > 0 Then result = result & "|Inherited:" & UXDefaultsStamp(0)
    UXDefaultsStamp = result
End Function
Public Function UXBlockValidated(ByVal changes As Object, ByVal kind As String) As Object
    Dim result As Object, key As Variant, v As Variant, choice As Variant, allowed As Boolean
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In changes.Keys
        v = changes(key)
        If Not IsNull(v) Then
            If UXFieldDate(key) And Len(v) > 0 Then
                If Not IsDate(v) Then Err.Raise vbObjectError + 4800, , UXFieldCaption(key, kind) & ": enter a valid date."
                v = Format$(CDate(v), "yyyy-mm-dd")
            End If
            If Len(UXFieldOptions(key)) > 0 Then
                allowed = False
                For Each choice In Split(UXFieldOptions(key), ";")
                    If SameFieldValue(choice, v) Then allowed = True
                Next
                If Not allowed Then Err.Raise vbObjectError + 4805, , UXFieldCaption(key, kind) & ": choose a listed value or clear it."
            End If
            If UXFieldLimit(key) > 0 And Len(v) > UXFieldLimit(key) Then Err.Raise vbObjectError + 4801, , UXFieldCaption(key, kind) & ": value is too long."
            If UXFieldDigits(key) > 0 And Len(v) > 0 Then
                If Len(v) <> UXFieldDigits(key) Or v Like "*[!0-9]*" Then Err.Raise vbObjectError + 4802, , UXFieldCaption(key, kind) & ": enter " & UXFieldDigits(key) & " digits."
            End If
        End If
        result(key) = v
    Next
    Set UXBlockValidated = result
End Function
Public Function UXGroupDefaultValue(ByVal groupID As Long, ByVal key As String) As Variant
    Dim id As Long
    id = Nz(DLookup("DefaultID", "FieldDefault", UXDefaultScopeWhere(groupID) & " AND FieldKey=" & QuoteSQL(key)), 0)
    If id > 0 Then AssertDefaultLink id
    UXGroupDefaultValue = DLookup("ValueText", "FieldDefault", UXDefaultScopeWhere(groupID) & " AND FieldKey=" & QuoteSQL(key))
End Function
Public Function UXBlockPendingValue(ByVal key As String) As Variant
    UXBlockPendingValue = UXBulkPending(key)
End Function
Public Function UXCloseBlocks() As Boolean
    DoCmd.Close acForm, "frmUXBulk", acSaveNo
    UXCloseBlocks = True
End Function
Public Function UXBlockViewport() As Boolean
    If Not CurrentProject.AllForms("frmUXBulk").IsLoaded Then Exit Function
    If IsNull(Forms!frmUXBulk!uxGroupContext) And Len(Nz(Forms!frmUXBulk.OpenArgs, "")) > 0 Then Forms!frmUXBulk!uxGroupContext = CLng(Forms!frmUXBulk.OpenArgs)
    Dim i As Long
    For i = 0 To 4
        If UXHasControl(Forms!frmUXBulk, "uxGroupTab" & i) Then Forms!frmUXBulk.Controls("uxGroupTab" & i).Visible = (Nz(Forms!frmUXBulk!uxGroupContext, 0) > 0 And i <> 3)
    Next
    Forms!frmUXBulk!uxAddGroupReport.Visible = (Nz(Forms!frmUXBulk!uxGroupContext, 0) > 0)
    UXResizeWorkspace "frmUXBulk"
    UXBlockResize
    UXBlockViewport = True
End Function
Public Sub UXBlockState(ByVal key As String, ByVal caption As String)
    ' Per-field inheritance/staging explanations have been removed.
End Sub
Public Function UXOpenGroupDefaults(Optional ByVal groupID As Long = -1) As Boolean
    UXOpenGroupDefaults = UXStartGroupBulk(True, False, groupID)
End Function
Public Function UXBlockMode(ByVal page As String) As Boolean
    UXBlockMode = UXImmediateBlockMode()
End Function
