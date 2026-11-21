Attribute VB_Name = "LayoutAuditHarness"
Option Compare Database
Option Explicit
' Layout audit: opens the main windows in realistic states and sizes and reports what a
' user would see as broken, without a screenshot: controls outside the window, controls
' covering each other, captions that do not fit their box (measured with Windows' own
' text layout at the screen's DPI), a window taller than the space it was laid out for,
' and buttons whose colors do not match whether they are available.
' Any finding fails, and so does missing one of the five defects planted at run time
' (proof that the checks still see what they are meant to).

Private Type AuditRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal index As Long) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeOut As Long, ByVal charSet As Long, ByVal outPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitchAndFamily As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal obj As LongPtr) As Long
Private Declare PtrSafe Function DrawTextW Lib "user32" (ByVal hdc As LongPtr, ByVal text As LongPtr, ByVal count As Long, ByRef rect As AuditRect, ByVal format As Long) As Long

Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CALCRECT As Long = &H400
Private Const DT_NOPREFIX As Long = &H800
Private Const LOGPIXELSY As Long = 90
Private Const SLACK_TWIPS As Long = 15

Private checks As Long, findings As Collection, stage As String

Private Sub Note(ByVal text As String)
    findings.Add stage & ": " & text
End Sub

' Rendered height in pixels of a caption wrapped to a width in pixels.
Private Function TextHeight(ByVal text As String, ByVal fontName As String, ByVal points As Single, ByVal weight As Long, ByVal italic As Boolean, ByVal widthPx As Long, ByVal dpi As Long, ByVal hdc As LongPtr, Optional ByVal wrap As Boolean = True) As Long
    Dim font As LongPtr, old As LongPtr, r As AuditRect
    font = CreateFontW(-CLng(points * dpi / 72), 0, 0, 0, weight, IIf(italic, 1, 0), 0, 0, 1, 0, 0, 5, 0, StrPtr(fontName))
    old = SelectObject(hdc, font)
    r.Right = widthPx
    DrawTextW hdc, StrPtr(text), Len(text), r, DT_CALCRECT Or DT_NOPREFIX Or IIf(wrap, DT_WORDBREAK, 0)
    SelectObject hdc, old: DeleteObject font
    If wrap Then TextHeight = r.Bottom Else TextHeight = r.Right
End Function

' What Access draws for a caption: && is one ampersand, a single & only marks the access key.
Private Function ShownCaption(ByVal caption As String) As String
    ShownCaption = Replace(Replace(Replace(caption, "&&", Chr$(1)), "&", ""), Chr$(1), "&")
End Function

' A control on a tab page that is not showing is not seen.
Private Function Showing(ByVal c As Access.Control) As Boolean
    Dim p As Object
    If Not c.Visible Then Exit Function
    On Error Resume Next
    Set p = c.Parent
    Do While Not p Is Nothing
        If TypeOf p Is Access.Form Then Exit Do
        If TypeOf p Is Access.Page Then
            If p.Parent.Value <> p.PageIndex Then Exit Function
        End If
        Set p = p.Parent
    Loop
    Showing = True
End Function

Private Function Kind(ByVal c As Access.Control) As String
    Select Case c.ControlType
        Case acLabel: If Len(Trim$(c.Caption)) > 0 And c.Width > 30 And c.Height > 30 Then Kind = "text"
        Case acCommandButton, acToggleButton: If Len(Trim$(c.Caption)) > 0 Then Kind = "button" Else Kind = "field"
        Case acTextBox, acComboBox, acListBox, acCheckBox, acOptionButton, acSubform: Kind = "field"
        Case acOptionGroup: Kind = "group"
    End Select
End Function

' The other control is this one's attached label or holds it (an option group).
Private Function Related(ByVal a As Access.Control, ByVal b As Access.Control) As Boolean
    On Error Resume Next
    If a.Parent.Name = b.Name Then Related = True: Exit Function
    If b.Parent.Name = a.Name Then Related = True: Exit Function
    If a.ControlType = acOptionGroup Or b.ControlType = acOptionGroup Then
        If a.Parent.Name = b.Parent.Name And (a.Parent.Name <> a.Parent.Parent.Name) Then Related = True
    End If
End Function

' Every finding for one open form (and the forms inside its subforms).
Public Function AuditForm(ByVal f As Access.Form, Optional ByVal layoutHeight As Long = 0, Optional ByVal prefix As String = "") As Long
    Dim c As Access.Control, d As Access.Control, list As New Collection, i As Long, j As Long, k As String
    Dim hdc As LongPtr, dpi As Long, need As Long, room As Long, text As String, sectionHeight As Long, total As Long
    Dim where As String, line As Long, extent() As Long
    where = prefix & f.Name
    hdc = GetDC(0): dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    On Error Resume Next
    For Each c In f.Controls
        If Len(Kind(c)) > 0 Then
            If Showing(c) Then list.Add c
        End If
    Next
    On Error GoTo 0
    ReDim extent(0 To list.Count)
    For i = 1 To list.Count
        Set c = list(i): extent(i) = c.Width
        If Kind(c) = "text" Then
            If c.TextAlign <= 1 And InStr(c.Caption, vbLf) = 0 Then
                need = TextHeight(ShownCaption(c.Caption), c.FontName, c.FontSize, c.FontWeight, c.FontItalic, 0, dpi, hdc, False) * 1440 \ dpi + c.LeftMargin + c.RightMargin + 60
                If need < extent(i) Then extent(i) = need
            End If
        End If
    Next
    For i = 1 To list.Count
        Set c = list(i): k = Kind(c)
        sectionHeight = f.Section(c.Section).Height
        checks = checks + 1
        If c.Left < 0 Or c.Top < 0 Or c.Left + c.Width > f.Width + SLACK_TWIPS Or c.Top + c.Height > sectionHeight + SLACK_TWIPS Then
            Note where & "." & c.Name & " is outside the window (" & c.Left & "," & c.Top & " " & c.Width & "x" & c.Height & "; form " & f.Width & "x" & sectionHeight & ")"
        End If
        If k = "text" Or k = "button" Then
            text = ShownCaption(c.Caption)
            line = TextHeight("Ag", c.FontName, c.FontSize, c.FontWeight, c.FontItalic, 10000, dpi, hdc)
            If k = "text" Then
                room = (c.Height - c.TopMargin - c.BottomMargin) * dpi \ 1440
                need = TextHeight(text, c.FontName, c.FontSize, c.FontWeight, c.FontItalic, (c.Width - c.LeftMargin - c.RightMargin) * dpi \ 1440 + 3, dpi, hdc)
            Else
                room = c.Height * dpi \ 1440 - 4
                need = TextHeight(text, c.FontName, c.FontSize, c.FontWeight, c.FontItalic, c.Width * dpi \ 1440 - 6, dpi, hdc)
            End If
            checks = checks + 1
            'One line is clipped only when the box is well under a line; wrapped text when it
            'runs a third of a line past the box (line spacing differs slightly from Access).
            If (need <= line And room * 4 < line * 3) Or (need > line And need > room + line \ 3) Then
                Note where & "." & c.Name & " text does not fit (" & need & " px needed, " & room & " px tall): """ & Left$(Replace(text, vbCrLf, " / "), 70) & """"
            End If
        End If
        For j = i + 1 To list.Count
            Set d = list(j)
            If c.Section = d.Section And Not Related(c, d) Then
                checks = checks + 1
                If c.Left + SLACK_TWIPS < d.Left + extent(j) And d.Left + SLACK_TWIPS < c.Left + extent(i) And c.Top + SLACK_TWIPS < d.Top + d.Height And d.Top + SLACK_TWIPS < c.Top + c.Height Then
                    Note where & "." & c.Name & " overlaps " & d.Name
                End If
            End If
        Next
    Next
    ReleaseDC 0, hdc
    If layoutHeight > 0 Then
        On Error Resume Next
        total = f.Section(acDetail).Height + f.Section(acHeader).Height + f.Section(acFooter).Height
        If total = 0 Then total = f.Section(acDetail).Height
        On Error GoTo 0
        checks = checks + 1
        If total > layoutHeight + SLACK_TWIPS Then Note where & " is " & (total - layoutHeight) & " twips taller than its window, so it scrolls"
    End If
    checks = checks + 1
    If Len(UXStaleButtons(f)) > 0 Then Note where & " buttons keep the wrong colors: " & UXStaleButtons(f)
    For Each c In f.Controls
        If c.ControlType = acSubform Then
            If Showing(c) And Len(c.SourceObject) > 0 Then AuditForm c.Form, 0, where & "/"
        End If
    Next
    AuditForm = findings.Count
End Function

Private Sub Scene(ByVal name As String, ByVal formName As String, Optional ByVal layoutHeight As Long = 0)
    stage = name
    AuditForm Forms(formName), layoutHeight
End Sub

' Size the editor's window; its own resize lays it out for the space it really has.
Private Sub SizeEditor(ByVal width As Long, ByVal height As Long)
    Forms!frmEditor.InsideWidth = width: Forms!frmEditor.InsideHeight = height
    UXFixedResize Forms!frmEditor.InsideHeight, Forms!frmEditor.InsideWidth
End Sub

Private Sub WorkspaceAt(ByVal formName As String, ByVal width As Long, ByVal height As Long)
    UXNavigate formName
    UXResizeWorkspace formName, width, height
    Scene formName & " " & width & "x" & height, formName
End Sub

' The auditor must catch four known faults planted at run time.
Private Function PlantedDefectsCaught() As String
    Dim f As Access.Form, before As Long, keep As Variant
    Set f = Forms!frmUXValidation
    stage = "planted"
    before = findings.Count
    keep = Array(f!cmdEvalman.Left, f!cmdEvalman.Top)
    f!cmdEvalman.Left = f!txtEvidence.Left + 200: f!cmdEvalman.Top = f!txtEvidence.Top + 200
    AuditForm f
    If Not FoundSince(before, "cmdEvalman overlaps txtEvidence") And Not FoundSince(before, "txtEvidence overlaps cmdEvalman") Then PlantedDefectsCaught = "a button over the evidence box was not flagged": Exit Function
    f!cmdEvalman.Left = keep(0): f!cmdEvalman.Top = keep(1)
    before = findings.Count
    keep = f!lblCount.Height
    f!lblCount.Height = 120
    AuditForm f
    If Not FoundSince(before, "lblCount text does not fit") Then PlantedDefectsCaught = "a caption too big for its label was not flagged": Exit Function
    f!lblCount.Height = keep
    before = findings.Count
    f!cmdOpen.BackColor = RGB(255, 0, 255)
    AuditForm f
    If Not FoundSince(before, "wrong colors: cmdOpen") Then PlantedDefectsCaught = "a miscolored button was not flagged": Exit Function
    UXEnable f!cmdOpen, f!cmdOpen.Enabled
    before = findings.Count
    AuditForm f, f.Section(acDetail).Height - 430
    If Not FoundSince(before, "taller than its window") Then PlantedDefectsCaught = "a window taller than its space was not flagged": Exit Function
    DropPlanted
End Function
'Remove the planted findings so only real ones are reported.
Private Sub DropPlanted()
    Dim i As Long
    For i = findings.Count To 1 Step -1
        If Left$(findings(i), 9) = "planted: " Then findings.Remove i
    Next
End Sub
Private Function FoundSince(ByVal before As Long, ByVal text As String) As Boolean
    Dim i As Long
    For i = before + 1 To findings.Count
        If InStr(findings(i), text) > 0 Then FoundSince = True: Exit Function
    Next
End Function

Public Function TestLayoutAudit(ByVal reportPath As String) As String
    Dim id As Long, other As Long, group As Long, member As Long, f As Access.Form, s As Access.Form, sizes As Variant, size As Variant
    Dim route As Variant, failure As String, file As Integer, item As Variant, page As Variant
    On Error GoTo Failed
    checks = 0: Set findings = New Collection
    stage = "setup"
    DoCmd.OpenForm "frmSetup": UXResizeWorkspace "frmSetup", 15000, 9000
    Scene "setup 15000x9000", "frmSetup"
    stage = "planted"
    Dim beforeSetup As Long, keepCaption As String, keepHeight As Long
    beforeSetup = findings.Count
    keepCaption = Forms!frmSetup!uxDataFileLocation.Caption: keepHeight = Forms!frmSetup!uxDataFileLocation.Height
    Forms!frmSetup!uxDataFileLocation.Caption = "Your work will be kept in a new data file, C:\Users\Example\Downloads\Swift NAVFIT data.accdb. This file stays the application, so a newer version can replace it without moving your data."
    Forms!frmSetup!uxDataFileLocation.Height = 650
    AuditForm Forms!frmSetup
    If Not FoundSince(beforeSetup, "uxDataFileLocation text does not fit") Then Err.Raise vbObjectError + 9600, , "Auditor missed the setup screen's original text overflow"
    DropPlanted
    Forms!frmSetup!uxDataFileLocation.Caption = keepCaption: Forms!frmSetup!uxDataFileLocation.Height = keepHeight
    DoCmd.Close acForm, "frmSetup", acSaveNo
    stage = "seed"
    ConfigureWorkspace "SYNTHETIC LAYOUT AUDIT COMMAND", "Layout auditor", "12345"
    member = SaveDirectoryMember(0, 0, "Jordan Layout-Example", "LAYOUT-EXAMPLE, JORDAN A", "IT2", "000000041")
    id = CommitDraft(0, "EVAL", "Layout audit report with a long local title", "Jordan Layout-Example", #1/1/2026#, #9/30/2026#, "", "Test")
    other = CommitDraft(0, "EVAL", "Second report", "Casey Second", #1/1/2026#, #9/30/2026#, "", "Test")
    group = CreateSummaryGroup("Layout audit summary group with a long name", "EVAL", "E5")
    AddGroupMember group, id, "E5", "": AddGroupMember group, other, "E5", ""
    'Workspace windows at a laptop size and a large monitor size.
    sizes = Array(Array(15000, 8500), Array(24000, 13500))
    For Each size In sizes
        For Each route In UXWorkspaceRoutes()
            WorkspaceAt Split(CStr(route), "|")(1), CLng(size(0)), CLng(size(1))
        Next
    Next
    'The report editor on each page, at both sizes; it must fit its window.
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor
    For Each size In sizes
        For Each page In Array(0, 1, 8, 9)
            UXFixedSelect CLng(page)
            SizeEditor CLng(size(0)), CLng(size(1))
            Scene "editor page " & page & " " & size(0) & "x" & size(1), "frmEditor", f.InsideHeight
        Next
    Next
    'With an invalid entry the error list shows above the page.
    UXFixedSelect 0
    Set s = f!uxFieldSurface.Form
    s!txtMemberSSN = "123": UXSurfaceEdited "MemberSSN"
    SizeEditor 15000, 9000
    Scene "editor with error list 15000x9000", "frmEditor", f.InsideHeight
    'Validation issues with a finding selected, at two widths.
    UXOpenValidation False
    Forms!frmUXValidation!lstIssues = Forms!frmUXValidation!lstIssues.ItemData(1): UXValidationSelection
    For Each size In Array(Array(12000, 9500), Array(17000, 9500))
        Forms!frmUXValidation.InsideWidth = CLng(size(0)): Forms!frmUXValidation.InsideHeight = CLng(size(1))
        UXValidationLayout
        Scene "validation " & size(0) & "x" & size(1), "frmUXValidation"
    Next
    failure = PlantedDefectsCaught()
    If Len(failure) > 0 Then Err.Raise vbObjectError + 9600, , "Auditor missed a planted defect: " & failure
    UXValidationClose
    'The Block guide beside the report, showing a block with examples and findings.
    stage = "block guide"
    Call UXOpenBlockGuide: Call UXBlockGuideFollow("ReadinessSummary"): Call UXLivePaint
    Scene "block guide", "frmUXBlockGuide"
    DoCmd.Close acForm, "frmUXBlockGuide", acSaveNo
    UXLeaveEditor "Discard"
    'Screens that show while code is blocked are checked as designed: nothing lays them out.
    For Each item In Array("frmEnableContent", "frmDataFileNotice")
        DoCmd.OpenForm CStr(item)
        Scene CStr(item), CStr(item)
        DoCmd.Close acForm, CStr(item), acSaveNo
    Next
    stage = "EVALMAN"
    UXOpenEvalmanSection EvalmanSectionForField("ReadinessSummary")
    Scene "EVALMAN", "frmUXEvalman"
    DoCmd.Close acForm, "frmUXEvalman", acSaveNo
    stage = "member history"
    UXDirectoryHistoryFor member
    Scene "member history", "frmUXMemberHistory"
    UXDirectoryClose "frmUXMemberHistory"
    stage = "group members"
    UXNavigate "frmUXGroups"
    Forms!frmUXGroups!lstGroups = group
    UXOpenGroup
    For Each size In sizes
        UXResizeWorkspace "frmUXComparison", CLng(size(0)), CLng(size(1))
        Scene "group members " & size(0) & "x" & size(1), "frmUXComparison"
    Next
    file = FreeFile
    Open reportPath For Output As #file
    For Each item In findings
        Print #file, item
    Next
    Close #file
    If findings.Count > 0 Then
        TestLayoutAudit = "FAIL: " & findings.Count & " layout findings (listed in " & reportPath & "): " & findings(1)
    Else
        TestLayoutAudit = "PASS: " & checks & " layout checks, no findings; all 5 planted defects caught"
    End If
    Exit Function
Failed:
    TestLayoutAudit = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
