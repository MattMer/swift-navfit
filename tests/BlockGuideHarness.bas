Attribute VB_Name = "BlockGuideHarness"
Option Compare Binary
Option Explicit
' Block guide: section choice for each field, sample entries, and the window following
' focus in the report editor without taking it.
Private n As Long

Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5310, , message
    n = n + 1
End Sub

Private Function TitleOf(ByVal section As Long) As String
    TitleOf = CStr(EvalmanHeader(section, 1))
End Function

' Move focus into a report field, opening the page that shows it.
Private Function FocusField(ByVal key As String) As Access.Control
    Dim f As Access.Form, s As Access.Form, c As Access.Control, page As Variant
    Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
    Set c = s.Controls(UXFieldControl(key))
    For Each page In Array(0, 2, 3, 4, 5, 6, 7)
        Call UXSelectSection(CLng(page))
        If c.Visible Then Exit For
    Next
    If Not c.Visible Then Err.Raise vbObjectError + 5311, , "No page shows " & key
    'A click in the report activates its window first; GotFocus follows.
    f.SetFocus: f!uxFieldSurface.SetFocus: c.SetFocus
    Set FocusField = c
End Function

Private Function Shown() As String
    Shown = Split(UXBlockGuideShown(), "|")(0)
End Function

Private Function HasFinding(ByVal key As String, ByRef message As String) As Boolean
    Dim item As Object
    If UXLiveFindings() Is Nothing Then Exit Function
    For Each item In UXLiveFindings()
        If CStr(item("Key")) = key Then message = CStr(item("Message")): HasFinding = True: Exit Function
    Next
End Function

Public Function TestBlockGuide() As String
    Dim kinds As Variant, kind As Variant, key As Variant, keys As Object, i As Long, s As Collection, item As Variant
    Dim id As Long, fitrep As Long, f As Access.Form, sf As Access.Form, c As Access.Control, g As Access.Form
    Dim fields As Object, result As Collection, finding As Object, message As String, text As String, found As Boolean, stage As String
    On Error GoTo Failed
    kinds = Array("EVAL", "CHIEFEVAL", "FITREP")

    stage = "titles"
    Check BlockGuideTitleRefs("Block 20: Physical Readiness") = "|G20|", "Single block refs"
    Check BlockGuideTitleRefs("Blocks 22-27: Reporting Senior") = "|G22|G23|G24|G25|G26|G27|", "Range refs"
    Check BlockGuideTitleRefs("Blocks 6 and 7: Uic/Ship/Station") = "|G6|G7|", "Pair refs"
    Check BlockGuideTitleRefs("FITREP Block 41: Comments on Performance") = "|E43|", "FITREP comments are EVAL block 43"
    Check BlockGuideTitleRefs("CHIEFEVAL Block 49: FITREP Block 46 Signature of Individual Evaluated") = "|E51|", "CHIEFEVAL signature is EVAL block 51"
    Check BlockGuideTitleRefs("EVAL Block 41: Career Recommendations") = "|E41|", "EVAL block 41 stays EVAL 41"
    Check BlockGuideTitleRefs("1-2. Report Blocks") = "", "Paragraph title has no block"

    stage = "sections"
    'Every field with an EVALMAN block opens that block first.
    Set keys = CreateObject("Scripting.Dictionary")
    For Each key In Split(BlockExampleKeys(), ",")
        keys(CStr(key)) = True
    Next
    For Each key In Array("OccasionPeriodic", "ReportConcurrent", "ComponentActive", "Trait33", "PromotionRecommendation", "RetentionYes", "StatementYes")
        keys(CStr(key)) = True
    Next
    For Each key In keys.Keys
        If EvalmanSectionForField(CStr(key)) >= 0 Then
            Set s = BlockGuideSections(CStr(key), "|")
            Check s.Count >= 1 And CLng(s(1)(0)) = EvalmanSectionForField(CStr(key)), "Main section wrong for " & key
            Check BlockGuideChapter(CLng(s(1)(0))) = 1, "Main section is not in chapter 1 for " & key
        End If
    Next
    found = False
    For Each item In BlockGuideSections("ReportingSeniorGrade", "|")
        If Left$(TitleOf(CLng(item(0))), 13) = "Blocks 22-27:" Then found = True
    Next
    Check found, "Block 23 does not list the Blocks 22-27 sections"
    Set s = BlockGuideSections("OfficialComments", "|")
    Check Left$(TitleOf(CLng(s(1)(0))), 14) = "EVAL Block 43:", "Comments do not open EVAL block 43"
    For Each item In s
        Check Left$(TitleOf(CLng(item(0))), 14) <> "EVAL Block 41:", "Comments list career recommendations"
    Next
    found = False
    For Each item In s
        If Left$(TitleOf(CLng(item(0))), 16) = "FITREP Block 41:" Then found = True
    Next
    Check found, "Comments do not list the FITREP block 41 sections"
    For Each item In Array(Array("|4|", 4), Array("|5|", 5), Array("|6|", 6))
        Set s = BlockGuideSections("ReadinessSummary", CStr(item(0)))
        Check s.Count > 2 And CBool(s(2)(1)) And BlockGuideChapter(CLng(s(2)(0))) = CLng(item(1)), "Block 20 does not list chapter " & item(1) & " first for " & item(0)
        For i = 3 To s.Count
            If CBool(s(i)(1)) Then Check BlockGuideChapter(CLng(s(i)(0))) = CLng(item(1)), "Wrong chapter marked as applying"
        Next
    Next
    Set s = BlockGuideSections("Title", "|")
    Check s.Count = 1 And CLng(s(1)(0)) = EvalmanBlocksSection(), "A field without a block does not show paragraph 1-2"
    Set s = BlockGuideSections("ShipStation", "|")
    found = False
    For Each item In s
        If Left$(TitleOf(CLng(item(0))), 8) = "Block 7:" Then found = True
    Next
    Check found, "Ship/station does not list the block 7 sections"

    stage = "examples"
    For Each key In Split(BlockExampleKeys(), ",")
        For Each kind In kinds
            For i = 0 To BlockExampleCount(CStr(key), CStr(kind)) - 1
                text = BlockExample(CStr(key), CStr(kind), i, 0)
                Check Len(text) > 0, "Empty example for " & key
                Check UXFieldLimit(CStr(key)) = 0 Or Len(text) <= UXFieldLimit(CStr(key)), "Example longer than the field: " & key & " " & text
                If BlockExample(CStr(key), CStr(kind), i, 2) = "EVALMAN" Then
                    Check EvalmanSectionByTitle(BlockExample(CStr(key), CStr(kind), i, 3)) >= 0, "EVALMAN example cites a missing section: " & key
                    Check InStr(1, EvalmanText(EvalmanSectionByTitle(BlockExample(CStr(key), CStr(kind), i, 3))), text, vbBinaryCompare) > 0, "EVALMAN example not in its section: " & key & " " & text
                End If
            Next
        Next
    Next
    Check BlockExampleCount("Trait35Comment3", "CHIEFEVAL") = BlockExampleCount("Trait35Comment1", "CHIEFEVAL"), "Trait comment boxes do not share examples"
    Check BlockExampleCount("RaterName", "FITREP") = 0 And BlockExampleCount("RaterName", "EVAL") > 0, "EVAL-only examples leak to FITREP"
    Check BlockExample("RankRate", "FITREP", 0, 0) <> BlockExample("RankRate", "EVAL", 0, 0), "Rank examples do not follow the report type"

    stage = "editor"
    ConfigureWorkspace "SYNTHETIC BLOCK GUIDE TEST", "Guide tester"
    id = CommitDraft(0, "EVAL", "Guide title", "Guide member", #11/16/2025#, #9/30/2026#, "", "Test", "Saved comments")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor: Set sf = f!uxFieldSurface.Form
    Check f!cmdBlockGuide.Caption = "Block guide" And f!cmdBlockGuide.OnClick = "=UXOpenBlockGuide()", "Block guide button missing"
    Check f!cmdBlockGuide.Visible And f!cmdBlockGuide.Left + f!cmdBlockGuide.Width <= f!cmdEvalman.Left, "Block guide button is not before EVALMAN"

    'Every report field calls the guide with its own key when it gets focus.
    For Each c In sf.Controls
        If Left$(c.Tag, 5) = "SURF:" Then
            Check c.OnGotFocus = "=UXBlockGuideFollow(""" & Mid$(c.Tag, 6) & """)", "Field does not follow focus: " & c.Name
            'Typed fields and block 21 have sample entries for each report type they appear on;
            'check boxes and fixed lists have none.
            key = Mid$(c.Tag, 6)
            For Each kind In kinds
                If UXFieldApplies(CStr(key), CStr(kind)) Then
                    If c.ControlType = acCheckBox Or (c.ControlType = acComboBox And key <> "BilletSubcategory") Then
                        Check BlockExampleCount(CStr(key), CStr(kind)) = 0, "A choice field has examples: " & key
                    ElseIf key <> "Title" And key <> "MemberName" And key <> "DraftNotes" Then
                        Check BlockExampleCount(CStr(key), CStr(kind)) > 0, "No examples for " & key & " on " & kind
                    End If
                End If
            Next
        End If
    Next

    'Single-line examples are entries Swift accepts.
    stage = "examples accepted"
    Set fields = EditorPendingFields()
    For Each key In Split(BlockExampleKeys(), ",")
        For Each kind In kinds
            For i = 0 To BlockExampleCount(CStr(key), CStr(kind)) - 1
                Set fields = EditorPendingFields(): fields("FormKind") = kind
                fields(CStr(key)) = BlockExample(CStr(key), CStr(kind), i, 0)
                For Each finding In ValidateLiveFields(fields, id)
                    Check Not (CStr(finding("Key")) = key And finding("Severity") = "Invalid entry"), "Example rejected: " & key & " " & fields(CStr(key)) & ": " & finding("Message")
                Next
            Next
        Next
    Next

    stage = "follow"
    Set c = FocusField("OfficialName")
    Check UXOpenBlockGuide() And CurrentProject.AllForms("frmUXBlockGuide").IsLoaded, "Guide did not open"
    Set g = Forms!frmUXBlockGuide
    Check g.PopUp And Not g.Modal, "Guide is not a modeless pop-up"
    Check sf.ActiveControl.Name = "txtOfficialName" And f.ActiveControl.Name = "uxFieldSurface", "Opening the guide moved the cursor out of the report"
    Check Not g!txtExamples.TabStop And Not g!txtGuide.TabStop, "Read-only guide text is a tab stop"
    Check Shown() = "OfficialName", "Guide did not open at the field last used: " & UXBlockGuideShown()
    'Focus moved while the guide's window is active raises no GotFocus; the editor's
    'timer catches up.
    f!uxFieldSurface.SetFocus: sf!txtRankRate.SetFocus
    Call UXLiveTimer
    Check Shown() = "RankRate", "The editor timer did not catch up with the field: " & UXBlockGuideShown()
    Set c = FocusField("BilletSubcategory")
    Check Shown() = "BilletSubcategory", "Guide did not follow focus to block 21: " & UXBlockGuideShown()
    Check sf.ActiveControl.Name = c.Name And f.ActiveControl.Name = "uxFieldSurface", "Guide took focus from the report"
    Check InStr(g!lblBlock.Caption, "21") > 0, "Block 21 caption missing: " & g!lblBlock.Caption
    Check Left$(g!lblSection.Caption, 9) = "Block 21:", "Block 21 section missing: " & g!lblSection.Caption
    Check InStr(Nz(g!txtExamples, ""), "NA") > 0, "Block 21 examples missing"
    Check InStr(Nz(g!txtGuide, ""), "Do not leave blank") > 0, "Block 21 instruction missing"

    stage = "findings"
    Call UXLivePaint()
    Check HasFinding("BilletSubcategory", message), "A blank block 21 has no finding"
    Check InStr(g!lblFindings.Caption, message) > 0, "Guide does not show the block 21 finding: " & g!lblFindings.Caption
    c.Value = "NA": Call UXSurfaceEdited("BilletSubcategory"): Call UXLivePaint()
    Check Not HasFinding("BilletSubcategory", message) And g!lblFindings.Caption = NO_FINDINGS, "Guide kept a cleared finding: " & g!lblFindings.Caption
    Set c = FocusField("PeriodEnd")
    c.Value = DateSerial(Year(Date) + 1, 12, 31): Call UXSurfaceEdited("PeriodEnd")
    Set c = FocusField("PeriodStart"): Call UXLivePaint()
    Check Shown() = "PeriodStart" And HasFinding("PeriodEnd", message), "No future-end finding"
    Check InStr(g!lblFindings.Caption, message) > 0, "A period end finding does not show on the period start"

    stage = "examples shown"
    Set c = FocusField("DateReported")
    Check InStr(Nz(g!txtExamples, ""), BlockExample("DateReported", "EVAL", 0, 0)) > 0, "Date reported example missing"
    Set c = FocusField("OccasionPeriodic")
    Check Shown() = "OccasionPeriodic" And IsNull(g!txtExamples), "A check box shows examples"
    Set c = FocusField("Trait36")
    Check InStr(Nz(g!txtGuide, ""), CatalogueTraitTitle("EVAL", 36)) > 0, "Trait title missing"

    stage = "situation"
    Set c = FocusField("ReportConcurrent")
    c.Value = True: Call UXSurfaceEdited("ReportConcurrent")
    Set c = FocusField("ReadinessSummary")
    Check InStr(g!lstOther.Column(1, 1), "This report: Concurrent") = 1, "Concurrent rules not listed first: " & g!lstOther.Column(1, 1)
    Check InStr(Nz(g!txtGuide, ""), "CONCURRENT REPORTS") > 0, "Concurrent rules not shown with block 20"
    g!lstOther = g!lstOther.ItemData(1): Call UXBlockGuideSelect
    Check BlockGuideChapter(CLng(Split(UXBlockGuideShown(), "|")(1))) = 4, "Choosing a section did not show it"
    Check UXBlockGuideEvalman() And UXEvalmanShown() = CLng(Split(UXBlockGuideShown(), "|")(1)), "Open in EVALMAN went elsewhere"
    DoCmd.Close acForm, "frmUXEvalman", acSaveNo

    stage = "close editor"
    Call UXLeaveEditor("Discard")
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded And CurrentProject.AllForms("frmUXBlockGuide").IsLoaded, "Guide closed with the editor"
    Check g!lblBlock.Caption = "Open a report to follow its blocks", "Guide kept a closed report: " & g!lblBlock.Caption
    Check UXBlockGuideFollow("RankRate") And g!lblBlock.Caption = "Open a report to follow its blocks", "Following with no editor failed"

    stage = "fitrep"
    fitrep = CommitDraft(0, "FITREP", "Guide FITREP", "Guide officer", #2/1/2026#, #1/31/2027#, "", "Test", "Saved comments")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(fitrep)
    Set f = Forms!frmEditor
    Set c = FocusField("OfficialComments")
    Check InStr(g!lblBlock.Caption, "41") > 0 And InStr(g!lblBlock.Caption, "FITREP") > 0, "FITREP comments are not block 41: " & g!lblBlock.Caption
    Check Left$(g!lblSection.Caption, 14) = "EVAL Block 43:", "FITREP comments opened the wrong section: " & g!lblSection.Caption
    Set c = FocusField("RankRate")
    Check InStr(Nz(g!txtExamples, ""), BlockExample("RankRate", "FITREP", 0, 0)) > 0, "FITREP grade example missing"
    Call UXLeaveEditor("Discard")
    DoCmd.Close acForm, "frmUXBlockGuide", acSaveNo
    Check Not UXBlockGuideFollow("RankRate"), "Closed guide still follows"

    TestBlockGuide = "PASS: " & n & " Block guide checks": Exit Function
Failed:
    TestBlockGuide = "FAIL after " & n & " at " & stage & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmUXEvalman", acSaveNo
    DoCmd.Close acForm, "frmUXBlockGuide", acSaveNo
    Call UXLeaveEditor("Discard")
End Function
