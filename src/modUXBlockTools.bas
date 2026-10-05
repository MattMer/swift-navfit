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
    Forms!frmUXBulk.Controls("uxValue" & key) = Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key).Value
    UXBlockEdit = UXBulkChanged(key)
    UXBlockState key, IIf(IsNull(Forms!frmUXBulk.Controls("uxValue" & key)), "Clear value staged", "Change staged")
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
        If UXHasControl(f, "value" & key) And UXFieldApplies(key, kind) Then
            flag = UXFieldFlag(key): multi = UXFieldMultiline(key): band = UXLayoutBand(CStr(key))
            size = UXLayoutWidth(CStr(key), usable, kind)
            If size < 2000 Then size = 2000
            If multi Then size = usable: band = CStr(key)
            If x > 150 And (band <> prior Or x + size > usable + 150) Then y = y + rowHeight: x = 150: rowHeight = 0
            With f.Controls("label" & key)
                .Caption = UXFieldCaption(key, kind): .Left = x: .Top = y: .Width = size - 500: .Height = 300: .Visible = True
            End With
            With f.Controls("value" & key)
                .Left = x: .Top = y + 320: .Width = size - 500: .Height = IIf(multi, 1800, 360): .Visible = True: .TabIndex = tabOrder
                tabOrder = tabOrder + 1
                If flag Then .Top = y: .Width = 280: .Height = 280
            End With
            If flag Then
                f.Controls("label" & key).Left = x + 360: f.Controls("label" & key).Width = size - 860
            End If
            With f.Controls("clear" & key)
                .Left = x + size - 450: .Top = y + IIf(flag, 0, 320): .Width = 450: .Height = 330: .Visible = True: .TabIndex = tabOrder
                tabOrder = tabOrder + 1
            End With
            With f.Controls("state" & key)
                .Left = x: .Top = y + IIf(flag, 320, IIf(multi, 2150, 700)): .Width = size: .Height = 260: .Visible = True
            End With
            If multi Then
                rowHeight = 2450
            ElseIf rowHeight < IIf(flag, 650, 1000) Then
                rowHeight = IIf(flag, 650, 1000)
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
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM FieldDefault WHERE ScopeKind='Group' AND ScopeID=" & groupID & " ORDER BY FieldKey", dbOpenSnapshot)
    Do While Not r.EOF
        result = result & r!FieldKey & ":" & r!VersionCounter & ":" & Nz(r!ValueText, "") & ":" & r!IsNullValue & "|"
        r.MoveNext
    Loop
    r.Close: UXDefaultsStamp = result
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
    UXGroupDefaultValue = DLookup("ValueText", "FieldDefault", "ScopeKind='Group' AND ScopeID=" & groupID & " AND FieldKey=" & QuoteSQL(key))
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
    UXResizeWorkspace "frmUXBulk"
    UXBlockResize
    UXBlockViewport = True
End Function
Public Sub UXBlockState(ByVal key As String, ByVal caption As String)
    Forms!frmUXBulk!uxBlocks.Form.Controls("state" & key).Caption = caption
End Sub
Public Function UXOpenGroupDefaults() As Boolean
    UXOpenGroupDefaults = UXStartGroupBulk(True)
End Function
Public Function UXBlockMode(ByVal page As String) As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXBulk: f!cmdPreview.SetFocus
    f!uxBlocks.Visible = (page <> "Preview"): f!lstPreview.Visible = (page = "Preview"): f!txtFullValues.Visible = (page = "Preview")
    Call UXBlockViewport: UXBlockMode = True
End Function

