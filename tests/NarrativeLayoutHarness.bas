Attribute VB_Name = "NarrativeLayoutHarness"
Option Explicit
Private checks As Long
Private referenceLastWords As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8061, , message
    checks = checks + 1
End Sub
Public Function TestNarrativeLayout() As String
    On Error GoTo Failed
    Dim kind As Variant, id As Long, f As Form, surface As Form, c As Control
    Dim width As Variant, value As String, revision As Long
    ConfigureWorkspace "SYNTHETIC NARRATIVE LAYOUT", "Tester"
    value = "Mixed CASE; quotes ""test"" and apostrophe '" & vbCrLf & String$(6000, "x")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = CommitDraft(0, CStr(kind), "Synthetic", "Synthetic", Null, Null, "", "Layout test", value)
        revision = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: f.TimerInterval = 0
        Check UXSelectSection(5), "Comments section failed"
        Set surface = f!uxFieldSurface.Form: Set c = surface!txtOfficialComments
        For Each width In Array(28000, 20000, 14000)
            Check UXResizeEditor(CLng(width), 11000), "Resize failed"
            Check c.Width = UXNarrativeWidth(CStr(kind)), "Narrative width changed with viewport"
            Check c.Height = UXNarrativeHeight(CStr(kind)), "Narrative height changed with viewport"
            Check c.FontName = "Courier New", "Font not locked"
            Check c.FontSize = UXNarrativeFontSize(CStr(kind)), "Point size differs"
            Check c.TextFormat = 0 And Not c.FontBold And Not c.FontItalic, "Rich styling remained"
            Check c.LeftMargin = 90 And c.RightMargin = 90, "Margins differ"
            Check surface.Controls("errOfficialComments").Top >= c.Top + c.Height, "Error overlaps narrative"
            Check c.Left + c.Width <= f!uxFieldSurface.Width - 300, "Narrative exceeds surface"
            Check c.Value = value, "Layout truncated overflow"
        Next
        ' Fixed sheet geometry: a narrow window scrolls the viewport; it never rewraps the narrative.
        Check UXResizeEditor(9000, 8000), "Narrow layout failed"
        Check c.Width = UXNarrativeWidth(CStr(kind)) And c.Height = UXNarrativeHeight(CStr(kind)), "Narrow viewport changed narrative geometry"
        Check f.ScrollBars = 3 And c.Left + c.Width <= f!uxFieldSurface.Width - 300,"Narrow viewport does not scroll the full narrative"
        Check InStr(surface!uxNarrativeCount.Caption, "Narrow view") = 0, "Narrow viewport claims changed wrapping"
        Check c.Value = value, "Narrow layout lost text"
        Check UXSelectSection(0), "Admin section failed"
        Check UXSelectSection(5), "Comments return failed"
        Check c.Value = value, "Section change lost text"
        Check DLookup("CurrentRevisionID", "Report", "ReportID=" & id) = revision, "Layout created revision"
        c = value & vbCrLf & "Pending tail"
        Check UXSurfaceEdited("OfficialComments"), "Visible edit failed"
        Check UXSaveEditor(), "Narrative save failed"
        Check DLookup("OfficialComments", "Report", "ReportID=" & id) = value & vbCrLf & "Pending tail", "Save truncated overflow"
        f.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check Forms!frmEditor!txtOfficialComments = value & vbCrLf & "Pending tail", "Reopen lost text"
        Forms!frmEditor.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    TestNarrativeLayout = "PASS: " & checks & " narrative font, geometry, viewport, overflow and save/reopen assertions"
    Exit Function
Failed:
    TestNarrativeLayout = "FAIL: " & kind & ": " & Err.Number & " " & Err.Description
End Function
' Reference wrap, word by word: a word stays on a line only when the line, the word and
' the space after it fit (the last word of a paragraph needs no space). Independent of
' the character-level LinesFor that PDF output and the meter share.
' Only for single-spaced words shorter than a line; it does not split long words.
' referenceLastWords receives the number of words on the last line.
Private Function ReferenceLines(ByVal text As String, ByVal width As Double, ByVal font As String, ByVal size As Double) As Long
    Dim paragraph As Variant, parts As Variant, line As String, i As Long, total As Long
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    If text = "" Then Exit Function
    If Right$(text, 1) = vbLf Then text = Left$(text, Len(text) - 1)
    For Each paragraph In Split(text, vbLf)
        total = total + 1: referenceLastWords = 0
        If Len(paragraph) > 0 Then
            parts = Split(paragraph, " "): line = parts(0): referenceLastWords = 1
            For i = 1 To UBound(parts)
                If PDFOutputTextWidth(line & " " & parts(i) & IIf(i < UBound(parts), " ", ""), font, size) > width - 1 Then
                    total = total + 1: line = parts(i): referenceLastWords = 1
                Else
                    line = line & " " & parts(i): referenceLastWords = referenceLastWords + 1
                End If
            Next
        End If
    Next
    ReferenceLines = total
End Function
Private Function Numbered(ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        If i > 1 Then Numbered = Numbered & vbCrLf
        Numbered = Numbered & "Line " & i
    Next
End Function
Private Function Words(ByVal count As Long) As String
    Dim names As Variant, i As Long
    names = Array("alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet", "kilo", "lima")
    For i = 0 To count - 1
        If i > 0 Then Words = Words & " "
        Words = Words & names(i Mod 12)
    Next
End Function
Private Sub CheckFit(ByVal kind As String, ByVal key As String, ByVal text As String, ByVal lines As Long, ByVal overflow As Long, ByVal oneWord As Boolean, ByVal label As String)
    Dim fit As Object
    Set fit = PDFFieldFit(kind, key, text)
    Check fit("Status") = "Measured", kind & " " & label & ": " & fit("Status") & " " & fit("Reason")
    Check fit("Lines") = lines, kind & " " & label & ": " & fit("Lines") & " lines, expected " & lines
    Check fit("Overflow") = overflow, kind & " " & label & ": overflow " & fit("Overflow") & ", expected " & overflow
    Check fit("OneWordLast") = oneWord, kind & " " & label & ": one-word last line " & fit("OneWordLast")
End Sub
' Line-fit meter: known line counts on all three forms, the editor caption and colour,
' the information-only finding, saving with overflow, and the 50 ms measurement budget.
Public Function TestNarrativeLineFit() As String
    On Error GoTo Failed
    Dim kind As Variant, key As Variant, fit As Object, capacity As Long, width As Double, font As String, size As Double
    Dim text As String, n As Long, id As Long, f As Form, surface As Form, fields As Object, item As Object
    Dim found As Boolean, counted As Long, started As Double, elapsed As Double, timing As String, stage As String, i As Long, limit As Long
    checks = 0
    ConfigureWorkspace "SYNTHETIC NARRATIVE LINE FIT", "Tester"
    For Each kind In Array("EVAL", "FITREP", "CHIEFEVAL")
        stage = kind & " geometry"
        Set fit = PDFFieldFit(CStr(kind), "OfficialComments", "")
        Check fit("Status") = "Measured", kind & ": printer metrics unavailable: " & fit("Reason")
        capacity = fit("Capacity")
        ' Line height is the font's own metric: Courier New 11 pt (about 11.8 pt lines) in a 201.5 pt
        ' box; Times New Roman 10 pt (11.5 pt lines) in 225.3 pt and 71.3 pt boxes.
        Select Case kind
            Case "EVAL": width = 539.5: font = "Courier New": size = 11: Check capacity = 17, kind & ": capacity " & capacity
            Case "FITREP": width = 536: font = "Times New Roman": size = 10: Check capacity = 19, kind & ": capacity " & capacity
            Case Else: width = 570.344: font = "Times New Roman": size = 10: Check capacity = 6, kind & ": capacity " & capacity
        End Select
        stage = kind & " known texts"
        CheckFit kind, "OfficialComments", "", 0, 0, False, "empty"
        CheckFit kind, "OfficialComments", Numbered(capacity), capacity, 0, False, "exactly full"
        CheckFit kind, "OfficialComments", Numbered(capacity + 1), capacity + 1, 1, False, "one line over"
        CheckFit kind, "OfficialComments", Numbered(capacity) & vbCrLf, capacity, 0, False, "full with trailing newline"
        CheckFit kind, "OfficialComments", Numbered(capacity) & vbCrLf & vbCrLf, capacity + 1, 1, False, "full with trailing blank line"
        CheckFit kind, "OfficialComments", "First" & vbCrLf & vbCrLf & vbCrLf & "Fourth line", 4, 0, False, "blank lines"
        CheckFit kind, "OfficialComments", "First line here" & vbCrLf & "End", 2, 0, True, "one-word last line"
        ' Wrapped text: grow word by word against the reference wrap.
        n = 1
        Do While ReferenceLines(Words(n + 1), width, font, size) = 1: n = n + 1: Loop
        CheckFit kind, "OfficialComments", Words(n), 1, 0, False, "one wrapped line"
        Check ReferenceLines(Words(n + 1), width, font, size) = 2, "Reference wrap"
        CheckFit kind, "OfficialComments", Words(n + 1), 2, 0, referenceLastWords = 1, "wrap to a second line"
        Check ReferenceLines(Words(n + 2), width, font, size) = 2, "Reference wrap"
        CheckFit kind, "OfficialComments", Words(n + 2), 2, 0, False, "wrap leaves two words"
        Do While ReferenceLines(Words(n + 1), width, font, size) <= capacity: n = n + 1: Loop
        Check ReferenceLines(Words(n), width, font, size) = capacity, "Reference wrap"
        CheckFit kind, "OfficialComments", Words(n), capacity, 0, referenceLastWords = 1, "wrapped exactly full"
        CheckFit kind, "OfficialComments", Words(n + 1), capacity + 1, 1, False, "wrapped one line over"
        text = Left$(Words(500), 2500)
        If Right$(text, 1) = " " Then text = Left$(text, 2499) & "x"
        Check Len(text) = 2500, "2,500-character text"
        n = ReferenceLines(text, width, font, size)
        CheckFit kind, "OfficialComments", text, n, IIf(n > capacity, n - capacity, 0), n <= capacity And referenceLastWords = 1, "2,500 characters"
        stage = kind & " timing"
        Set fit = PDFFieldFit(CStr(kind), "OfficialComments", text)
        started = Timer
        For i = 1 To 10: Set fit = PDFFieldFit(CStr(kind), "OfficialComments", text): Next
        elapsed = (Timer - started) * 100
        Check elapsed < 50, kind & ": a 2,500-character narrative took " & Format$(elapsed, "0.0") & " ms"
        timing = timing & " " & kind & "=" & Format$(elapsed, "0.0") & "ms"
        stage = kind & " other boxes"
        For Each key In Array("CommandEmployment", "DutyDescription")
            Set fit = PDFFieldFit(CStr(kind), CStr(key), Numbered(10))
            Check fit("Status") = "Measured" And fit("Capacity") >= 2 And fit("Overflow") = 10 - fit("Capacity"), kind & " " & key & " fit"
        Next
        Check PDFFieldFit(CStr(kind), "NoSuchField", "x")("Status") = "None", kind & ": a field without a print box measured"
        If kind = "CHIEFEVAL" Then
            CheckFit kind, "Trait33Comment1", "Short comment", 1, 0, False, "Chief comment"
            CheckFit kind, "Trait39Comment3", Numbered(3), 3, 1, False, "Chief comment over"
        End If
        Check PDFFieldFit("UNKNOWN", "OfficialComments", "x")("Status") = "None", "Unknown report type"
        stage = kind & " finding"
        limit = IIf(kind = "CHIEFEVAL", 6, 18)
        Check NAVFITFieldFit(CStr(kind), "OfficialComments", "")("Capacity") = limit, kind & ": NAVFIT98A comments limit"
        id = CommitDraft(0, CStr(kind), "Synthetic", "Synthetic", Null, Null, "", "Line fit test", Numbered(limit + 1))
        Set fields = SavedFields(id)
        If kind = "CHIEFEVAL" Then fields("Trait35Comment2") = Numbered(3)
        found = False: counted = 0
        For Each item In ValidateLiveFields(fields, id)
            If item("Code") = "APP-NARRATIVE-FIT" Then
                Check item("Severity") = "Information", kind & ": finding is not information-only"
                If item("Key") = "OfficialComments" Then
                    Check InStr(item("Message"), "Narrative is longer than NAVFIT98A's block: " & limit + 1 & " of " & limit & " lines in NAVFIT98A: 1 line over its limit.") = 1, kind & ": finding text " & item("Message")
                    found = True
                End If
                counted = counted + 1
            End If
        Next
        Check found And counted = IIf(kind = "CHIEFEVAL", 2, 1), kind & ": narrative findings " & counted
        fields("OfficialComments") = Numbered(limit)
        If kind = "CHIEFEVAL" Then fields("Trait35Comment2") = "Short"
        For Each item In ValidateLiveFields(fields, id)
            Check item("Code") <> "APP-NARRATIVE-FIT", kind & ": a fitting narrative has a finding"
        Next
        stage = kind & " editor meter"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: f.TimerInterval = 0
        Check UXSelectSection(5), "Comments section failed"
        Set surface = f!uxFieldSurface.Form
        Check UXNarrativeStatus(), "Status refresh failed"
        Check InStr(surface!uxNarrativeCount.Caption, limit + 1 & " of " & limit & " lines in NAVFIT98A: 1 line over its limit | ") = 1, kind & ": caption " & surface!uxNarrativeCount.Caption
        Check surface!uxNarrativeCount.BackColor = UXColor("Error") And surface!uxNarrativeCount.ForeColor = UXColor("ErrorInk"), kind & ": overflow is not red"
        surface!txtOfficialComments = Numbered(limit)
        Check UXSurfaceEdited("OfficialComments"), "Visible edit failed"
        Check UXNarrativeStatus(), "Status refresh failed"
        Check InStr(surface!uxNarrativeCount.Caption, limit & " of " & limit & " lines in NAVFIT98A: last line | ") = 1, kind & ": caption " & surface!uxNarrativeCount.Caption
        Check surface!uxNarrativeCount.BackColor = UXColor("Warning"), kind & ": last line is not amber"
        surface!txtOfficialComments = "First line here" & vbCrLf & "End"
        Check UXSurfaceEdited("OfficialComments"), "Visible edit failed"
        Check UXNarrativeStatus(), "Status refresh failed"
        Check InStr(surface!uxNarrativeCount.Caption, "2 of " & limit & " lines in NAVFIT98A; last line holds one word | ") = 1, kind & ": caption " & surface!uxNarrativeCount.Caption
        Check surface!uxNarrativeCount.BackStyle = 0 And surface!uxNarrativeCount.ForeColor = UXColor("Ink"), kind & ": fitting text is coloured"
        surface!txtOfficialComments = text
        Check UXSurfaceEdited("OfficialComments"), "Visible edit failed"
        Check UXLivePaint(), "Validation paint failed"
        counted = 0
        For Each item In ValidateLiveFields(EditorPendingFields(), id)
            If item("Severity") <> "Information" Then counted = counted + 1
        Next
        Check InStr(f!uxDirty.Caption, "| " & counted & " issues") > 0, kind & ": information counted as an issue: " & f!uxDirty.Caption
        Check UXSaveEditor(), kind & ": overflowing narrative did not save"
        Check DLookup("OfficialComments", "Report", "ReportID=" & id) = text, kind & ": saved narrative differs"
        f.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    PDFFieldFitRelease
    TestNarrativeLineFit = "PASS: " & checks & " line-fit count, meter, finding and save assertions; 2,500-character measurement" & timing
    Exit Function
Failed:
    TestNarrativeLineFit = "FAIL: " & stage & ": " & Err.Number & " " & Err.Description
End Function
' NAVFIT98A profile against the System.Drawing reference in test-narratives. Each manifest
' row: kind|key|lines|limit|one-word last line (1, 0, or -1 unchecked)|UTF-16 text file|case.
Public Function TestNavfitReference(ByVal manifest As String) As String
    On Error GoTo Failed
    Dim stream As Object, row As Variant, parts As Variant, text As String, fit As Object, label As String
    Dim started As Double, elapsed As Double, slowest As Double, rows As Long
    checks = 0
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile manifest: text = stream.ReadText: stream.Close
    Set fit = NAVFITFieldFit("EVAL", "OfficialComments", "Warm up GDI+ and the printer context")
    For Each row In Split(Replace(text, vbCr, ""), vbLf)
        If Len(row) > 0 Then
            parts = Split(row, "|"): label = parts(0) & " " & parts(1) & " " & parts(6): rows = rows + 1
            Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "unicode": stream.Open
            stream.LoadFromFile parts(5): text = stream.ReadText: stream.Close
            started = Timer
            Set fit = NAVFITFieldFit(CStr(parts(0)), CStr(parts(1)), text)
            elapsed = (Timer - started) * 1000: If elapsed > slowest Then slowest = elapsed
            Check fit("Status") = "Measured", label & ": " & fit("Status") & " " & fit("Reason")
            Check fit("Lines") = CLng(parts(2)), label & ": " & fit("Lines") & " lines; reference " & parts(2)
            Check fit("Capacity") = CLng(parts(3)), label & ": limit " & fit("Capacity")
            Check fit("Overflow") = IIf(CLng(parts(2)) > CLng(parts(3)), CLng(parts(2)) - CLng(parts(3)), 0), label & ": overflow " & fit("Overflow")
            If parts(4) <> "-1" Then Check fit("OneWordLast") = (parts(4) = "1"), label & ": one-word last line " & fit("OneWordLast")
        End If
    Next
    Check rows > 0, "No reference cases"
    Check slowest < 50, "Slowest NAVFIT98A measurement took " & Format$(slowest, "0") & " ms"
    Check NAVFITFieldFit("EVAL", "Trait33Comment1", "x")("Status") = "None", "EVAL has no trait comment box"
    Check NAVFITFieldFit("UNKNOWN", "OfficialComments", "x")("Status") = "None", "Unknown report type"
    PDFFieldFitRelease
    TestNavfitReference = "PASS: " & checks & " NAVFIT98A line-count assertions match the System.Drawing reference; slowest " & Format$(slowest, "0") & " ms"
    Exit Function
Failed:
    TestNavfitReference = "FAIL: " & label & ": " & Err.Number & " " & Err.Description
End Function
