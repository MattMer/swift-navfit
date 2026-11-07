Attribute VB_Name = "modUXBlockTools"
Option Explicit
Private blockSection As Long
Private laidOut As String, laidForm As Access.Form
Public Function UXBlockRefresh() As Boolean
    Dim f As Access.Form, key As Variant
    Set f = Forms!frmUXBulk!uxBlocks.Form
    For Each key In UXSharedKeys()
        f.Controls("value" & key) = UXSurfaceShown(CStr(key), Forms!frmUXBulk.Controls("uxValue" & key).Value)
    Next
    UXBlockRefresh = True
End Function
Public Function UXBlockEdit(ByVal key As String) As Boolean
    Forms!frmUXBulk.Controls("uxValue" & key) = UXBlockEntry(key, Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key).Value)
    UXBlockEdit = UXConfirmBlockChange(key)
End Function
Public Function UXBlockClear(ByVal key As String) As Boolean
    Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key) = Null
    UXBlockClear = UXBlockEdit(key)
End Function
'Entries follow the report editor: block 20's PFA date is a month (MM/YY).
Private Function UXBlockEntry(ByVal key As String, ByVal value As Variant) As Variant
    If key = "ReadinessDate" Then UXBlockEntry = UXReadinessDateEntry(value) Else UXBlockEntry = UXDateEntry(value)
End Function
'Defaults sit where the report's Front Page puts the same blocks, with the same
'captions and fonts. Rows the defaults do not use are closed up, and the page is
'scaled to the viewing width when it is narrower than the report.
Public Function UXBlockLayout(ByVal section As Long) As Boolean
    Dim f As Access.Form, c As Access.Control, row As Variant, parts As Variant, key As String, kind As String, usable As Long, defaultKeys As Object, item As Variant
    Dim rows As New Collection, tops As Object, ordered() As Long, i As Long, j As Long, swap As Long, compact As Object, y As Long, last As Long, lastHeight As Long
    Dim ratio As Double, x As Long, w As Long, h As Long, top As Long, tabOrder As Long, bottom As Long
    Set f = Forms!frmUXBulk!uxBlocks.Form: blockSection = section
    kind = Nz(Forms!frmUXBulk!uxKind, "EVAL"): usable = Forms!frmUXBulk!uxBlocks.Width - 700
    f!anchor.SetFocus
    For Each c In f.Controls
        If c.Name <> "anchor" Then c.Visible = False: c.Left = 100: c.Top = 100: c.Width = 200
    Next
    Set defaultKeys = CreateObject("Scripting.Dictionary")
    For Each item In UXSharedKeys(): defaultKeys(CStr(item)) = True: Next
    Set tops = CreateObject("Scripting.Dictionary")
    For Each row In ReportPageRows(kind, 0)
        parts = Split(CStr(row), "|"): key = parts(0)
        If defaultKeys.Exists(key) And UXHasControl(f, "value" & key) And UXFieldApplies(key, kind) And Not UXFieldLegacy(key) Then
            rows.Add parts
            If Not tops.Exists(CStr(parts(2))) Then tops.Add CStr(parts(2)), CLng(parts(4)) Else If CLng(parts(4)) > tops(CStr(parts(2))) Then tops(CStr(parts(2))) = CLng(parts(4))
        End If
    Next
    If rows.Count = 0 Then UXBlockLayout = True: Exit Function
    ReDim ordered(0 To tops.Count - 1)
    i = 0
    For Each item In tops.Keys: ordered(i) = CLng(item): i = i + 1: Next
    For i = 0 To UBound(ordered) - 1
        For j = i + 1 To UBound(ordered)
            If ordered(j) < ordered(i) Then swap = ordered(i): ordered(i) = ordered(j): ordered(j) = swap
        Next
    Next
    Set compact = CreateObject("Scripting.Dictionary")
    y = 150: last = ordered(0)
    For i = 0 To UBound(ordered)
        If i > 0 Then
            lastHeight = tops(CStr(ordered(i - 1))) + 700
            If ordered(i) - last > lastHeight Then y = y + lastHeight Else y = y + ordered(i) - last
        End If
        compact.Add CStr(ordered(i)), y: last = ordered(i)
    Next
    ratio = 1
    If usable < 11800 Then ratio = usable / 11800
    If ratio < 0.6 Then ratio = 0.6
    For Each parts In rows
        key = parts(0): top = compact(CStr(parts(2)))
        x = 150 + CLng(CLng(parts(1)) * ratio): w = CLng(CLng(parts(3)) * ratio): h = CLng(parts(4))
        With f.Controls("label" & key)
            .Caption = Replace(UXFieldCaption(key, kind), "&", "&&"): .FontSize = 8
            .Left = x: .Top = top: .Width = w: .Height = 220: .Visible = True
            If w < Len(UXFieldCaption(key, kind)) * 72 Then .Top = top - 180: .Height = 400
        End With
        Set c = f.Controls("value" & key)
        c.Left = x: c.Top = top + 230: c.Width = w: c.Height = h: c.Visible = True
        c.TabIndex = tabOrder: tabOrder = tabOrder + 1
        If c.ControlType <> acCheckBox Then c.FontName = "Segoe UI": c.FontSize = 10
        If UXFieldFlag(key) Then
            c.Left = x: c.Top = top + 40: c.Width = 280: c.Height = 280
            With f.Controls("label" & key): .Left = x + 350: .Top = top: .Width = w - 350: .Height = 430: End With
        End If
        If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
    Next
    f.Width = usable + 300: f.Section(acDetail).Height = bottom + 300
    Set laidForm = f: laidOut = section & "|" & kind & "|" & usable
    UXBlockLayout = True
End Function
' A resize reflows the fields only when the usable width (or section or report
' type) changed; every reflow hides and repositions every field.
Public Function UXBlockResize() As Boolean
    If Not CurrentProject.AllForms("frmUXBulk").IsLoaded Then Exit Function
    If laidForm Is Forms!frmUXBulk!uxBlocks.Form And laidOut = blockSection & "|" & Nz(Forms!frmUXBulk!uxKind, "EVAL") & "|" & (Forms!frmUXBulk!uxBlocks.Width - 700) Then UXBlockResize = True: Exit Function
    UXBlockResize = UXBlockLayout(blockSection)
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
                If Not IsDate(v) Then Err.Raise vbObjectError + 4800, "modUXBlockTools.UXBlockValidated", UXFieldCaption(key, kind) & ": enter a valid date."
                v = Format$(CDate(v), "yyyy-mm-dd")
            End If
            If Len(UXFieldOptions(key)) > 0 Then
                allowed = False
                For Each choice In Split(UXFieldOptions(key), ";")
                    If SameFieldValue(choice, v) Then allowed = True
                Next
                If Not allowed Then Err.Raise vbObjectError + 4805, "modUXBlockTools.UXBlockValidated", UXFieldCaption(key, kind) & ": choose a listed value or clear it."
            End If
            If UXFieldLimit(key) > 0 And Len(v) > UXFieldLimit(key) Then Err.Raise vbObjectError + 4801, "modUXBlockTools.UXBlockValidated", UXFieldCaption(key, kind) & ": value is too long."
            If UXFieldDigits(key) > 0 And Len(v) > 0 Then
                If Len(v) <> UXFieldDigits(key) Or v Like "*[!0-9]*" Then Err.Raise vbObjectError + 4802, "modUXBlockTools.UXBlockValidated", UXFieldCaption(key, kind) & ": enter " & UXFieldDigits(key) & " digits."
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
