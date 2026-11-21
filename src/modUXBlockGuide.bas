Attribute VB_Name = "modUXBlockGuide"
Option Compare Database
Option Explicit
' Block guide: a window beside the report editor that follows the field with focus. Each
' report field calls UXBlockGuideFollow with its key when it gets focus; the guide shows
' the block's name for this report type, its current findings, sample entries
' (modBlockExamples), the EVALMAN paragraph 1-2 instruction, the same block's rules in
' the chapters that apply to this report (Concurrent, Ops Cdr, NOB, Inactive), and the
' block's other chapter sections. It never takes focus from the editor.
Private shownKey As String, shownKind As String, shownSignature As String
Private shownSection As Long
Private Const GUIDE As String = "frmUXBlockGuide"
Public Const NO_FINDINGS As String = "No findings for this block."

Private Function GuideOpen() As Boolean
    GuideOpen = CurrentProject.AllForms(GUIDE).IsLoaded
End Function

Private Function EditorOpen() As Boolean
    EditorOpen = CurrentProject.AllForms("frmEditor").IsLoaded
End Function

Private Function Quoted(ByVal value As String) As String
    Quoted = """" & Replace(value, """", """""") & """"
End Function

' The block numbers an EVALMAN title covers, as |G20| for every form or |E43| for the EVAL
' numbering. FITREP and CHIEFEVAL titles that repeat an EVAL block are given its EVAL number.
Public Function BlockGuideTitleRefs(ByVal title As String) As String
    Dim head As String, form As String, rest As String, a As Long, b As Long, n As Long, out As String, p As Long
    p = InStr(title, ":"): If p > 0 Then head = Left$(title, p - 1) Else head = title
    If head Like "EVAL Block*" Then form = "E": head = Mid$(head, 6)
    If head Like "FITREP Block*" Then form = "F": head = Mid$(head, 8)
    If head Like "CHIEFEVAL Block*" Then form = "C": head = Mid$(head, 11)
    If head Like "Blocks *" Then
        rest = Mid$(head, 8)
    ElseIf head Like "Block *" Then
        rest = Mid$(head, 7)
    Else
        Exit Function
    End If
    rest = Trim$(rest)
    If InStr(rest, "-") > 0 Then
        a = Val(rest): b = Val(Mid$(rest, InStr(rest, "-") + 1))
    ElseIf InStr(rest, " and ") > 0 Then
        a = Val(rest): b = Val(Mid$(rest, InStr(rest, " and ") + 5))
        If b <> a + 1 Then out = "|" & BlockRef(form, a) & "|" & BlockRef(form, b) & "|": BlockGuideTitleRefs = out: Exit Function
    Else
        a = Val(rest): b = a
    End If
    If a <= 0 Or b < a Then Exit Function
    out = "|"
    For n = a To b
        out = out & BlockRef(form, n) & "|"
    Next
    BlockGuideTitleRefs = out
End Function

Private Function BlockRef(ByVal form As String, ByVal n As Long) As String
    Select Case form & n
        Case "F41", "C40": BlockRef = "E43"
        Case "F46", "C49": BlockRef = "E51"
        Case Else
            If form = "" Then BlockRef = "G" & n Else BlockRef = form & n
    End Select
End Function

' The references a field matches: its block's title, plus block 7 for ship/station,
' whose instructions sit inside the block 6 section.
Private Function FieldRefs(ByVal key As String) As String
    FieldRefs = BlockGuideTitleRefs(EvalmanTitleForField(key))
    If key = "ShipStation" Then FieldRefs = "|G7|"
End Function

' The chapter a section belongs to, from its printed page (4-5 is chapter 4), or 0.
Public Function BlockGuideChapter(ByVal section As Long) As Long
    Dim page As String
    page = CStr(EvalmanHeader(section, 2))
    If InStr(page, "-") > 1 Then BlockGuideChapter = Val(Left$(page, InStr(page, "-") - 1))
End Function

Private Function ChapterName(ByVal chapter As Long) As String
    Select Case chapter
        Case 1: ChapterName = "All reports"
        Case 3: ChapterName = "Regular reports"
        Case 4: ChapterName = "Concurrent reports"
        Case 5: ChapterName = "Operational Commander reports"
        Case 6: ChapterName = "Not Observed (NOB) reports"
        Case 7: ChapterName = "Civilian or foreign activities"
        Case 8: ChapterName = "Duty under instruction (DUINS)"
        Case 9: ChapterName = "Inactive duty Reservists"
        Case 10: ChapterName = "Reservists on temporary active duty"
        Case Else: ChapterName = "Chapter " & chapter
    End Select
End Function

' The chapters that apply to this report, as |4|6| (Concurrent, Ops Cdr, NOB, Inactive).
Public Function BlockGuideSituation(ByVal fields As Object) As String
    Dim out As String
    out = "|"
    If FieldIs(fields, "ReportConcurrent", "Yes") Then out = out & "4|"
    If FieldIs(fields, "ReportOperationalCommander", "Yes") Then out = out & "5|"
    If FieldIs(fields, "ObservationStatus", "NOB") Then out = out & "6|"
    If FieldIs(fields, "ComponentInactive", "Yes") Then out = out & "9|"
    BlockGuideSituation = out
End Function

Private Function FieldIs(ByVal fields As Object, ByVal key As String, ByVal value As String) As Boolean
    If fields Is Nothing Then Exit Function
    If fields.Exists(key) Then FieldIs = (Trim$(Nz(fields(key), "")) = value)
End Function

' Sections for a field: item 1 is the main (paragraph 1-2) section, then the sections in
' chapters that apply to this report, then the rest. Each item is Array(section, applies).
Public Function BlockGuideSections(ByVal key As String, ByVal situation As String) As Collection
    Dim out As New Collection, main As Long, refs As String, i As Long, r As Variant, others As New Collection, matched As Boolean
    Set BlockGuideSections = out
    main = EvalmanSectionForField(key)
    If main < 0 Then
        If EvalmanBlocksSection() >= 0 Then out.Add Array(EvalmanBlocksSection(), False)
        Exit Function
    End If
    out.Add Array(main, False)
    refs = FieldRefs(key)
    If Len(refs) = 0 Then Exit Function
    For i = 0 To EvalmanCount() - 1
        If i <> main Then
            matched = False
            For Each r In Split(Mid$(BlockGuideTitleRefs(CStr(EvalmanHeader(i, 1))), 2), "|")
                If Len(r) > 0 And InStr(refs, "|" & r & "|") > 0 Then matched = True: Exit For
            Next
            If matched Then
                If InStr(situation, "|" & BlockGuideChapter(i) & "|") > 0 Then out.Add Array(i, True) Else others.Add Array(i, False)
            End If
        End If
    Next
    For Each r In others
        out.Add r
    Next
End Function

Private Function SectionCaption(ByVal section As Long) As String
    SectionCaption = EvalmanHeader(section, 1) & IIf(Len(EvalmanHeader(section, 2)) > 0, ", page " & EvalmanHeader(section, 2), "")
End Function

' Examples for a field on this report type, one per paragraph: the entry, then its note.
Public Function BlockGuideExamplesText(ByVal key As String, ByVal kind As String) As String
    Dim i As Long, out As String, note As String
    For i = 0 To BlockExampleCount(key, kind) - 1
        If Len(out) > 0 Then out = out & vbCrLf
        out = out & BlockExample(key, kind, i, 0)
        If BlockExample(key, kind, i, 2) = "EVALMAN" Then out = out & "   (EVALMAN's example)"
        note = BlockExample(key, kind, i, 1)
        If Len(note) > 0 Then out = out & vbCrLf & "      " & note
    Next
    BlockGuideExamplesText = out
End Function

' Findings for the shown block: a finding on any field of the same block counts.
Private Function FindingsText(ByVal key As String) As String
    Dim result As Collection, item As Object, title As String, out As String, n As Long, extra As Long
    Set result = UXLiveFindings()
    If result Is Nothing Then Exit Function
    title = EvalmanTitleForField(key)
    For Each item In result
        If CStr(item("Key")) = key Or (Len(title) > 0 And EvalmanTitleForField(CStr(item("Key"))) = title) Then
            n = n + 1
            If n <= 3 Then
                out = out & IIf(Len(out) > 0, vbCrLf, "") & item("Severity") & ": " & item("Message")
            Else
                extra = extra + 1
            End If
        End If
    Next
    If extra > 0 Then out = out & vbCrLf & extra & " more in Validation checks."
    FindingsText = Left$(out, 1000)
End Function

Private Function TraitText(ByVal key As String, ByVal kind As String, ByVal fields As Object) As String
    Dim block As Long, mark As String
    If Not (key Like "Trait3#" Or key Like "Trait3#Comment#") Then Exit Function
    block = Val(Mid$(key, 6, 2))
    TraitText = block & ". " & CatalogueTraitTitle(kind, block) & vbCrLf & CatalogueTraitDescription(kind, block, "")
    If fields Is Nothing Then Exit Function
    If fields.Exists("Trait" & block) Then mark = Trim$(Nz(fields("Trait" & block), ""))
    If Len(mark) > 0 Then TraitText = TraitText & vbCrLf & vbCrLf & "Selected mark: " & mark & vbCrLf & CatalogueTraitDescription(kind, block, mark)
End Function

Private Function ReportKind() As String
    ReportKind = "EVAL"
    If EditorOpen() Then ReportKind = Nz(Forms!frmEditor!cboKind, "EVAL")
End Function

Private Function PendingFields() As Object
    On Error Resume Next
    If EditorOpen() Then Set PendingFields = EditorPendingFields()
End Function

' Open the guide beside the editor at the field last used there, or at paragraph 1-2.
Public Function UXOpenBlockGuide() As Boolean
    Dim key As String, c As Access.Control
    If Not GuideOpen() Then DoCmd.OpenForm GUIDE
    On Error Resume Next
    If EditorOpen() Then
        Set c = Forms!frmEditor!uxFieldSurface.Form.ActiveControl
        If Not c Is Nothing Then
            If Left$(c.Tag, 5) = "SURF:" Then key = Mid$(c.Tag, 6)
        End If
    End If
    shownSignature = ""
    Call ShowField(key)
    'Opening the guide leaves the cursor where it was in the report.
    If Not c Is Nothing Then Forms!frmEditor.SetFocus: Forms!frmEditor!uxFieldSurface.SetFocus: c.SetFocus
    On Error GoTo 0
    UXOpenBlockGuide = True
End Function

Public Function UXLoadBlockGuide() As Boolean
    shownKey = "": shownKind = "": shownSignature = "": shownSection = -1
    With Forms(GUIDE)
        !lblSource.Caption = EvalmanSource() & " text; the PDF is authoritative."
        .Caption = "Block guide"
    End With
    Call UXBlockGuideLayout
    UXLoadBlockGuide = True
End Function

' The focus hook on every report field. Errors are ignored: a repaint must never interrupt
' moving between fields.
Public Function UXBlockGuideFollow(ByVal key As String) As Boolean
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    Call ShowField(key)
    UXBlockGuideFollow = True
End Function

' The editor's timer: Access does not raise a field's GotFocus when focus moves while
' another window is active, so the guide catches up with the field that has focus.
Public Function UXBlockGuideTick() As Boolean
    Dim c As Access.Control, key As String
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    Set c = Forms!frmEditor!uxFieldSurface.Form.ActiveControl
    If c Is Nothing Then Exit Function
    If Left$(c.Tag, 5) <> "SURF:" Then Exit Function
    key = Mid$(c.Tag, 6)
    If key <> shownKey Then Call ShowField(key)
    UXBlockGuideTick = True
End Function

' Repaint when the report type, its situation or (for a trait) the mark changed, too.
Public Function UXBlockGuideRefresh() As Boolean
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    Call ShowField(shownKey)
    UXBlockGuideRefresh = True
End Function

Private Sub ShowField(ByVal key As String)
    Dim f As Access.Form, kind As String, fields As Object, situation As String, signature As String
    Dim sections As Collection, s As Variant, rows As String, text As String, trait As String, caption As String, applies As Boolean
    Set f = Forms(GUIDE)
    If Not EditorOpen() Then ShowClosed: Exit Sub
    kind = ReportKind(): Set fields = PendingFields(): situation = BlockGuideSituation(fields)
    signature = key & "|" & kind & "|" & situation
    If key Like "Trait3#" Or key Like "Trait3#Comment#" Then
        If Not fields Is Nothing Then
            If fields.Exists("Trait" & Val(Mid$(key, 6, 2))) Then signature = signature & "|" & Nz(fields("Trait" & Val(Mid$(key, 6, 2))), "")
        End If
    End If
    If signature = shownSignature Then UXBlockGuideFindings: Exit Sub
    shownKey = key: shownKind = kind: shownSignature = signature
    If Len(key) > 0 And Len(EvalmanTitleForField(key)) > 0 Then
        caption = UXFieldCaption(key, kind) & " (" & kind & ")"
    ElseIf Len(key) > 0 Then
        caption = UXFieldCaption(key, kind) & ": no EVALMAN block for this field"
    Else
        caption = "Click a field in the report"
    End If
    f!lblBlock.Caption = caption
    'Check boxes and fixed lists have no sample entries; the box is left empty.
    text = ""
    If Len(key) > 0 Then text = BlockGuideExamplesText(key, kind)
    If Len(text) > 0 Then f!txtExamples = text Else f!txtExamples = Null
    f!lblExampleHead.Caption = IIf(Len(text) > 0, "For example", "For example: none for this kind of field")
    Set sections = BlockGuideSections(key, situation)
    If sections.Count = 0 Then f!txtGuide = Null: f!lstOther.RowSource = "": f!lblSection.Caption = "": Exit Sub
    shownSection = CLng(sections(1)(0))
    trait = TraitText(key, kind, fields)
    text = IIf(Len(trait) > 0, trait & vbCrLf & vbCrLf, "") & EvalmanText(shownSection)
    rows = ";" & shownSection & ";" & Quoted(ChapterName(BlockGuideChapter(shownSection)) & ": " & SectionCaption(shownSection))
    For Each s In sections
        If CLng(s(0)) <> shownSection Then
            applies = CBool(s(1))
            If applies Then text = text & vbCrLf & vbCrLf & "=== " & UCase$(ChapterName(BlockGuideChapter(CLng(s(0))))) & ": " & SectionCaption(CLng(s(0))) & " ===" & vbCrLf & EvalmanText(CLng(s(0)))
            rows = rows & ";" & s(0) & ";" & Quoted(IIf(applies, "This report: ", "") & ChapterName(BlockGuideChapter(CLng(s(0)))) & ": " & SectionCaption(CLng(s(0))))
        End If
    Next
    f!lblSection.Caption = SectionCaption(shownSection)
    f!txtGuide = text
    f!lstOther.RowSource = Mid$(rows, 2)
    f!lblOther.Caption = IIf(sections.Count > 1, "This block in other chapters (" & sections.Count - 1 & ")", "This block in other chapters (none)")
    Call UXBlockGuideFindings
End Sub

Private Sub ShowClosed()
    Dim f As Access.Form
    If shownSignature = "closed" Then Exit Sub
    Set f = Forms(GUIDE)
    f!lblBlock.Caption = "Open a report to follow its blocks"
    f!lblSection.Caption = "": f!lblFindings.Caption = "": f!lblExampleHead.Caption = "For example"
    f!txtExamples = Null: f!txtGuide = Null: f!lstOther.RowSource = ""
    shownKey = "": shownSignature = "closed": shownSection = -1
End Sub

' The editor is closing (its Close event, while it is still loaded): the guide stays open
' and waits for the next report.
Public Function UXBlockGuideEditorClosed() As Boolean
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    ShowClosed
    UXBlockGuideEditorClosed = True
End Function

' Live findings for the shown block; called after each live validation paint.
Public Function UXBlockGuideFindings() As Boolean
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    Dim text As String
    Dim role As String
    If Len(shownKey) > 0 And EditorOpen() Then
        text = FindingsText(shownKey)
        If Len(text) = 0 Then text = NO_FINDINGS: role = "Muted"
    End If
    If Len(role) = 0 Then role = IIf(InStr(text, "Invalid entry") > 0 Or InStr(text, "Policy issue") > 0, "ErrorInk", "WarningInk")
    With Forms(GUIDE)!lblFindings
        If .Caption <> text Then .Caption = text
        UXVisualSet Forms(GUIDE)!lblFindings, "ForeColor", UXColor(role)
    End With
    UXBlockGuideFindings = True
End Function

' A chosen section replaces the text until focus moves to another field.
Public Function UXBlockGuideSelect() As Boolean
    Dim section As Long
    section = Val(Nz(Forms(GUIDE)!lstOther, -1))
    If section < 0 Or section >= EvalmanCount() Then Exit Function
    shownSection = section
    Forms(GUIDE)!lblSection.Caption = SectionCaption(section)
    Forms(GUIDE)!txtGuide = EvalmanText(section)
    shownSignature = ""
    UXBlockGuideSelect = True
End Function

Public Function UXBlockGuideEvalman() As Boolean
    If shownSection < 0 Then shownSection = EvalmanBlocksSection()
    UXBlockGuideEvalman = UXOpenEvalmanSection(shownSection)
End Function

' The field and EVALMAN section shown, for tests: "key|section".
Public Function UXBlockGuideShown() As String
    UXBlockGuideShown = shownKey & "|" & shownSection
End Function

' Controls stretch with the window; the EVALMAN text takes the height that is left.
Public Function UXBlockGuideLayout() As Boolean
    Dim f As Access.Form, w As Long, h As Long, c As Variant, bottom As Long
    If Not GuideOpen() Then Exit Function
    On Error Resume Next
    Set f = Forms(GUIDE)
    w = f.InsideWidth - 400: If w < 4800 Then w = 4800
    h = f.InsideHeight: If h < 7600 Then h = 7600
    For Each c In Array("lblBlock", "lblSection", "lblFindings", "lblExampleHead", "txtExamples", "lblGuideHead", "txtGuide", "lblOther", "lstOther", "lblSource")
        If f.Controls(c).Width <> w Then f.Controls(c).Width = w
    Next
    bottom = h - 250
    f!lblSource.Top = bottom - f!lblSource.Height
    f!cmdClose.Top = f!lblSource.Top - 80 - f!cmdClose.Height
    f!cmdEvalman.Top = f!cmdClose.Top: f!cmdWording.Top = f!cmdClose.Top
    f!cmdClose.Left = 200 + w - f!cmdClose.Width
    f!lstOther.Top = f!cmdClose.Top - 120 - f!lstOther.Height
    f!lblOther.Top = f!lstOther.Top - f!lblOther.Height - 30
    f!txtGuide.Height = f!lblOther.Top - 100 - f!txtGuide.Top
    f.Section(acDetail).Height = h
    UXBlockGuideLayout = True
End Function
