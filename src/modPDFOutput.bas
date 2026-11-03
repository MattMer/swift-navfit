Attribute VB_Name = "modPDFOutput"
Option Compare Binary
Option Explicit

'One bounded preview snapshot. Resetting VBA invalidates it; it never auto-applies.
Private job As Object
Private printPages As Collection
Private cachePath As String, outputFailure As String
Private exporting As Boolean
Private previewFolder As String
'Text slots and background the open print report last showed ("*" = unknown).
Private shownSlots As Long, shownBackground As String
Private Type PDFExtent
    Width As Long
    Height As Long
End Type
Private Type PDFFontMetrics
    Height As Long
    Ascent As Long
    Descent As Long
    InternalLeading As Long
    ExternalLeading As Long
    Remaining(0 To 39) As Byte
End Type
'Report.FontName/TextWidth are unavailable during planning (Access 2455).
'Use Windows printer GDI metrics; no UI automation or external renderer.
'Point sizing: learn.microsoft.com/windows/win32/api/wingdi/nf-wingdi-createfontw
Private Declare PtrSafe Function PDFCreateDC Lib "gdi32" Alias "CreateDCW" (ByVal driver As LongPtr, ByVal device As LongPtr, ByVal output As LongPtr, ByVal init As LongPtr) As LongPtr
Private Declare PtrSafe Function PDFDeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function PDFDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal dc As LongPtr, ByVal index As Long) As Long
Private Declare PtrSafe Function PDFCreateFont Lib "gdi32" Alias "CreateFontW" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strike As Long, ByVal charset As Long, ByVal precision As Long, ByVal clip As Long, ByVal quality As Long, ByVal pitch As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function PDFSelectObject Lib "gdi32" Alias "SelectObject" (ByVal dc As LongPtr, ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function PDFDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function PDFTextExtent Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal dc As LongPtr, ByVal value As LongPtr, ByVal count As Long, ByRef result As PDFExtent) As Long
Private Declare PtrSafe Function PDFTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal dc As LongPtr, ByRef result As PDFFontMetrics) As Long
Private Declare PtrSafe Function PDFTextFace Lib "gdi32" Alias "GetTextFaceW" (ByVal dc As LongPtr, ByVal count As Long, ByVal buffer As LongPtr) As Long
Private metricDC As LongPtr, metricFont As LongPtr, metricOldFont As LongPtr
'NAVFIT98A prints narratives with .NET Graphics.DrawString (GDI+) into fixed rectangles.
'GdipMeasureString is the same layout engine. learn.microsoft.com/windows/win32/api/gdiplusflat/
Private Type PDFRectF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type
Private Type PDFGdiplusInput
    Version As Long
    Callback As LongPtr
    NoThread As Long
    NoCodecs As Long
End Type
Private Declare PtrSafe Function PDFGdiplusStartup Lib "gdiplus" Alias "GdiplusStartup" (ByRef token As LongPtr, ByRef options As PDFGdiplusInput, ByVal output As LongPtr) As Long
Private Declare PtrSafe Sub PDFGdiplusShutdown Lib "gdiplus" Alias "GdiplusShutdown" (ByVal token As LongPtr)
Private Declare PtrSafe Function PDFGdipCreateFromHDC Lib "gdiplus" Alias "GdipCreateFromHDC" (ByVal dc As LongPtr, ByRef graphics As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipDeleteGraphics Lib "gdiplus" Alias "GdipDeleteGraphics" (ByVal graphics As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipSetPageUnit Lib "gdiplus" Alias "GdipSetPageUnit" (ByVal graphics As LongPtr, ByVal unit As Long) As Long
Private Declare PtrSafe Function PDFGdipCreateFontFamily Lib "gdiplus" Alias "GdipCreateFontFamilyFromName" (ByVal name As LongPtr, ByVal collection As LongPtr, ByRef family As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipDeleteFontFamily Lib "gdiplus" Alias "GdipDeleteFontFamily" (ByVal family As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipCreateFont Lib "gdiplus" Alias "GdipCreateFont" (ByVal family As LongPtr, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, ByRef font As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipDeleteFont Lib "gdiplus" Alias "GdipDeleteFont" (ByVal font As LongPtr) As Long
Private Declare PtrSafe Function PDFGdipMeasureString Lib "gdiplus" Alias "GdipMeasureString" (ByVal graphics As LongPtr, ByVal text As LongPtr, ByVal length As Long, ByVal font As LongPtr, ByRef layout As PDFRectF, ByVal format As LongPtr, ByRef bounds As PDFRectF, ByRef fitted As Long, ByRef lines As Long) As Long
Private gdipToken As LongPtr
Private metricDpiX As Long, metricDpiY As Long, metricKey As String
Public Const PDF_OUTPUT_SLOTS As Long = 340

Public Function PDFOutputBusy() As Boolean
    PDFOutputBusy = exporting
End Function
Public Function PDFOutputJob() As Object
    If job Is Nothing Then Err.Raise vbObjectError + 9100, "modPDFOutput.PDFOutputJob", "Refresh the PDF preview; its captured values are no longer available."
    Set PDFOutputJob = job
End Function
Public Function PDFOutputText(ByVal fields As Object, ByVal key As String) As String
    If fields.Exists(key) Then PDFOutputText = Nz(fields(key), "")
End Function
Public Function PDFOutputSafePart(ByVal value As String, Optional ByVal limit As Long = 48) As String
    Dim i As Long, n As Long, c As String, result As String, masked As String
    'SSN-shaped and ten-digit identifiers never reach a file name.
    i = 1
    Do While i <= Len(value)
        n = IdentifierLength(value, i)
        If n > 0 Then
            masked = masked & "ID": i = i + n
        Else
            masked = masked & Mid$(value, i, 1): i = i + 1
        End If
    Loop
    value = masked
    For i = 1 To Len(value)
        c = Mid$(value, i, 1)
        If AscW(c) >= 32 Or AscW(c) < 0 Then
            If InStr(1, "<>:""/\|?*", c, vbBinaryCompare) > 0 Or c = " " Then c = "_"
            result = result & c
        End If
    Next
    result = Left$(Trim$(result), limit)
    Do While Len(result) > 0
        If Right$(result, 1) <> "." And Right$(result, 1) <> " " Then Exit Do
        result = Left$(result, Len(result) - 1)
    Loop
    If result = "" Then result = "Unspecified"
    If ReservedDeviceName(result) Then result = "_" & result
    PDFOutputSafePart = result
End Function
' Length of an identifier starting at i, else 0: ten digits, or 3-2-4 digits with
' an optional space or hyphen between groups (the former VBScript pattern
' "[0-9]{10}|[0-9]{3}[- ]?[0-9]{2}[- ]?[0-9]{4}").
Private Function IdentifierLength(ByVal text As String, ByVal i As Long) As Long
    Dim k As Long, part As Variant
    If DigitsAt(text, i, 10) Then IdentifierLength = 10: Exit Function
    k = i
    For Each part In Array(3, 2, 4)
        If part <> 3 And (Mid$(text, k, 1) = "-" Or Mid$(text, k, 1) = " ") Then
            If DigitsAt(text, k + 1, CLng(part)) Then k = k + 1
        End If
        If Not DigitsAt(text, k, CLng(part)) Then Exit Function
        k = k + part
    Next
    IdentifierLength = k - i
End Function
Private Function DigitsAt(ByVal text As String, ByVal i As Long, ByVal count As Long) As Boolean
    Dim k As Long
    If i < 1 Or i + count - 1 > Len(text) Then Exit Function
    For k = i To i + count - 1
        If Not IsDigitAt(text, k) Then Exit Function
    Next
    DigitsAt = True
End Function
' Windows reserved device names (CON, PRN, AUX, NUL, COM1-9, LPT1-9), with or without an extension.
Private Function ReservedDeviceName(ByVal name As String) As Boolean
    Dim base As String, dot As Long
    dot = InStr(name, ".")
    If dot > 0 Then base = UCase$(Left$(name, dot - 1)) Else base = UCase$(name)
    Select Case base
        Case "CON", "PRN", "AUX", "NUL": ReservedDeviceName = True
        Case Else
            If Len(base) = 4 And (Left$(base, 3) = "COM" Or Left$(base, 3) = "LPT") Then ReservedDeviceName = (Right$(base, 1) >= "1" And Right$(base, 1) <= "9")
    End Select
End Function
Private Function ShortIdentity(ByVal identity As String) As String
    ShortIdentity = Left$(Replace(Replace(Replace(identity, "{", ""), "}", ""), "-", ""), 8)
End Function
Public Function PDFOutputFileName(ByVal title As String, ByVal row As Object) As String
    Dim ending As String
    ending = PDFOutputText(row("Fields"), "PeriodEnd")
    If ending = "" Then ending = "Undated"
    PDFOutputFileName = PDFOutputSafePart(title) & "__" & PDFOutputSafePart(row("Caption")) & "__" & PDFOutputSafePart(row("Kind"), 12) & "__" & PDFOutputSafePart(ending, 12) & "__r" & row("Revision") & "_" & ShortIdentity(row("Guid")) & ".pdf"
End Function
Private Function CaptureRow(ByVal id As Long, ByVal takePending As Boolean) As Object
    Dim row As Object, fields As Object, revision As Long, name As String, last As String, first As String, middle As String, parts As Variant, source As String
    Set row = CreateObject("Scripting.Dictionary")
    If id > 0 Then
        revision = ReportHeadRevision(id): Set fields = RevisionFields(revision)
        row.Add "Guid", CStr(MajorScalar("SELECT ReportGuid FROM Report WHERE ReportID=" & id))
        source = "Saved revision " & revision
    Else
        row.Add "Guid", NewIdentity(): source = "Unsaved draft"
    End If
    If takePending And CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = id Then Set fields = EditorPendingFields(): source = "Captured editor draft"
    End If
    If fields Is Nothing Then Err.Raise vbObjectError + 9101, "modPDFOutput.CaptureRow", "Open the draft or choose a saved report."
    If takePending And id > 0 Then
        If UXGridPDFSnapshot(id, fields) Then
            If source = "Captured editor draft" Then source = "Editor draft + pending group cells (group cells win)" Else source = "Captured group-table draft"
        End If
    End If
    name = PDFOutputText(fields, "OfficialName"): If name = "" Then name = PDFOutputText(fields, "MemberName")
    If name = "" Then name = "Unnamed report"
    last = PDFOutputText(fields, "MemberLastName"): first = PDFOutputText(fields, "MemberFirstName"): middle = PDFOutputText(fields, "MemberMiddleInitial")
    If last = "" Then
        parts = Split(name, ","): last = Trim$(parts(0))
        If UBound(parts) > 0 And first = "" Then first = Trim$(parts(1))
    End If
    row.Add "ID", id: row.Add "Revision", revision: row.Add "Fields", fields
    row.Add "Caption", name: row.Add "Last", last: row.Add "First", first: row.Add "Middle", middle
    row.Add "Kind", PDFOutputText(fields, "FormKind"): row.Add "Source", source
    row.Add "Findings", FindingText(fields, id)
    Set CaptureRow = row
End Function
Private Function FindingText(ByVal fields As Object, ByVal id As Long) As String
    Dim findings As Collection, item As Object, result As String
    On Error GoTo Unavailable
    Set findings = ValidateLiveFields(fields, id)
    For Each item In findings
        result = result & item("Severity") & ": " & item("Message") & vbCrLf
    Next
    FindingText = result: Exit Function
Unavailable:
    FindingText = "Validation could not complete. PDF generation remains available."
End Function
Public Function PDFOutputCompare(ByVal a As Object, ByVal b As Object) As Long
    Dim key As Variant, result As Long
    For Each key In Array("Last", "First", "Middle", "Guid")
        result = StrComp(CStr(a(key)), CStr(b(key)), vbTextCompare)
        If result <> 0 Then PDFOutputCompare = result: Exit Function
    Next
End Function
Private Sub AddSorted(ByVal rows As Collection, ByVal row As Object)
    Dim i As Long
    For i = 1 To rows.Count
        If PDFOutputCompare(row, rows(i)) < 0 Then rows.Add row, , i: Exit Sub
    Next
    rows.Add row
End Sub
Public Sub PDFOutputCapture(ByVal groupID As Long, ByVal reportID As Long, ByVal takePending As Boolean)
    Dim rows As New Collection, files As New Collection, stats As New Collection, r As DAO.Recordset, row As Object, context As Long
    If exporting Then Err.Raise vbObjectError + 9102, "modPDFOutput.PDFOutputCapture", "Wait for the current PDF export to finish."
    PDFOutputEndPreview
    Set job = CreateObject("Scripting.Dictionary")
    job.Add "ID", NewIdentity(): job.Add "Captured", Now(): job.Add "GroupID", groupID
    job.Add "Title", "Ungrouped": job.Add "Paygrade", "": job.Add "GroupKind", ""
    job.Add "ISICID", GroupISIC(groupID): job.Add "OnTime", ""
    If groupID > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT * FROM SummaryGroup WHERE GroupID=" & groupID, dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 9103, "modPDFOutput.PDFOutputCapture", "The summary group is unavailable."
        job("Title") = Nz(r!Title, "Summary group"): job("Paygrade") = Nz(r!Paygrade, ""): job("GroupKind") = Nz(r!FormKind, ""): r.Close
        Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID, dbOpenSnapshot)
        Do While Not r.EOF
            Set row = CaptureRow(CLng(r!ReportID), takePending): AddSorted rows, row: r.MoveNext
        Loop
        r.Close
    Else
        Set row = CaptureRow(reportID, takePending): rows.Add row
    End If
    If groupID > 0 Then
        context = groupID: Set stats = rows
    ElseIf reportID > 0 Then
        If DCount("*", "GroupMember", "ReportID=" & reportID) = 1 Then
            context = CLng(DLookup("GroupID", "GroupMember", "ReportID=" & reportID))
            job("Title") = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & context), "Summary group")
            Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & context, dbOpenSnapshot)
            Do While Not r.EOF
                Set row = CaptureRow(CLng(r!ReportID), takePending): stats.Add row: r.MoveNext
            Loop
            r.Close
        End If
    End If
    job.Add "Reports", rows: job.Add "Files", files: job.Add "Folder", ""
    job.Add "StatisticsGroupID", context: job.Add "StatisticsRows", stats
End Sub
'Summary letter choices made on the output screen: the ISIC and whether the report
'was submitted on time ("Yes", "No" or blank).
Public Sub PDFOutputSetLetter(ByVal isic As Long, ByVal onTime As String)
    If job Is Nothing Then Exit Sub
    job("ISICID") = isic: job("OnTime") = onTime
End Sub
Public Function PDFOutputFolderPreference() As String
    Dim identity As String
    identity = Nz(MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1"), "")
    If identity <> "" Then PDFOutputFolderPreference = GetSetting("SwiftNAVFIT", "PDF-" & identity, "Folder", "")
End Function
Public Sub PDFOutputRememberFolder(ByVal folder As String)
    Dim identity As String
    identity = Nz(MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1"), "")
    If identity <> "" Then SaveSetting "SwiftNAVFIT", "PDF-" & identity, "Folder", folder
End Sub
Public Function PDFOutputReportOpen(ByVal report As Access.Report) As Boolean
    Dim sql As String
    On Error GoTo Failed
    If job Is Nothing Then Err.Raise vbObjectError + 9100, "modPDFOutput.PDFOutputReportOpen", "PDF preview expired. Refresh it before generating."
    If cachePath = "" Or printPages Is Nothing Then Err.Raise 5, "modPDFOutput.PDFOutputReportOpen", "PDF page cache is unavailable."
    sql = "SELECT * FROM PDFPage IN '" & Replace(cachePath, "'", "''") & "' ORDER BY PageNo"
    report.RecordSource = sql
    'A new report instance starts from its design, so the first page resets every control.
    shownSlots = PDF_OUTPUT_SLOTS: shownBackground = "*"
    PDFOutputReportOpen = True: Exit Function
Failed:
    outputFailure = "Open print report: " & Err.Number & " " & Err.Description
End Function
Public Function PDFOutputFormat(ByVal report As Access.Report) As Boolean
    Dim page As Object, slot As Variant, c As Access.Control, i As Long, name As Variant, used As Long, stage As String
    On Error GoTo Failed
    stage = "record lookup"
    If report.HasData = 0 Then Exit Function
    Set page = printPages(CLng(report!txtPageNo))
    stage = "reset controls"
    'Hide only what the previous page showed; resetting all 340 slots on every
    'page dominated packet time.
    If shownBackground = "*" Then
        For Each name In PDFPrintBackgrounds()
            report.Controls("bg_" & name).Visible = False
        Next
    ElseIf shownBackground <> "" And shownBackground <> page("Background") Then
        report.Controls("bg_" & shownBackground).Visible = False
    End If
    For i = page("Slots").Count + 1 To shownSlots: report.Controls("v" & Format$(i, "000")).Visible = False: Next
    If page("Background") <> "" Then report.Controls("bg_" & page("Background")).Visible = True
    shownBackground = page("Background")
    If shownSlots < page("Slots").Count Then shownSlots = page("Slots").Count
    For Each slot In page("Slots")
        used = used + 1: Set c = report.Controls("v" & Format$(used, "000"))
        stage = "slot " & used & " on page " & report.CurrentRecord
        c.Value = slot(0)
        c.Left = 0: c.Top = 0: c.Width = CLng(slot(3) * 20): c.Height = CLng(slot(4) * 20)
        c.Left = CLng(slot(1) * 20): c.Top = CLng(slot(2) * 20)
        c.FontName = slot(5): c.FontSize = slot(6): c.TextAlign = slot(7)
        c.FontBold = slot(9)
        c.BackStyle = IIf(slot(8), 1, 0): c.BackColor = vbWhite: c.Visible = True
    Next
    shownSlots = used
    PDFOutputFormat = True: Exit Function
Failed:
    shownSlots = PDF_OUTPUT_SLOTS: shownBackground = "*"
    If outputFailure = "" Then outputFailure = "Format PDF page (" & stage & "): " & Err.Number & " " & Err.Description
End Function
Private Sub ClosePrintReport()
    If CurrentProject.AllReports("rptPDFOutput").IsLoaded Then DoCmd.Close acReport, "rptPDFOutput", acSaveNo
    CloseMetricContext
End Sub
Public Sub PDFOutputEndPreview()
    Dim fs As Object
    ClosePrintReport
    If previewFolder = "" Then Exit Sub
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(cachePath) Then fs.DeleteFile cachePath
    RmDir previewFolder
    previewFolder = "": cachePath = ""
End Sub
Public Function PDFOutputReportClosed() As Boolean
    If CurrentProject.AllForms("frmPDFOutput").IsLoaded Then Forms!frmPDFOutput.Visible = True
    PDFOutputReportClosed = True
End Function
Public Sub PDFOutputPreviewPages(ByVal root As String, ByVal includeSummary As Boolean)
    Dim rows As Collection, row As Object, fs As Object
    PDFOutputEndPreview
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(root) Then Err.Raise 76, "modPDFOutput.PDFOutputPreviewPages", "Choose an available protected export folder."
    Set rows = job("Reports"): Set printPages = New Collection
    If includeSummary And job("GroupID") > 0 Then PlanSummary rows
    For Each row In rows: PlanReport row, rows: Next
    ClosePrintReport
    If printPages.Count = 0 Then Err.Raise 5, "modPDFOutput.PDFOutputPreviewPages", "No pages are available."
    previewFolder = fs.BuildPath(root, "PDF_Preview_" & ShortIdentity(NewIdentity()) & ".partial")
    fs.CreateFolder previewFolder: cachePath = previewFolder & "\pages.accdb"
    Call WritePageCache: outputFailure = ""
    DoCmd.OpenReport "rptPDFOutput", acViewPreview
    If outputFailure <> "" Then Err.Raise vbObjectError + 9108, "modPDFOutput.PDFOutputPreviewPages", outputFailure
End Sub
Private Function NewPage(ByVal background As String) As Object
    Dim page As Object, slots As New Collection
    Set page = CreateObject("Scripting.Dictionary"): page.Add "Background", background: page.Add "Slots", slots
    printPages.Add page: Set NewPage = page
End Function
Private Sub AddSlot(ByVal page As Object, ByVal value As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, Optional ByVal font As String = "Times New Roman", Optional ByVal size As Double = 10, Optional ByVal align As Long = 1, Optional ByVal opaque As Boolean = False, Optional ByVal bold As Boolean = False)
    If value = "" Then Exit Sub
    If page("Slots").Count >= PDF_OUTPUT_SLOTS Then Err.Raise vbObjectError + 9105, "modPDFOutput.AddSlot", "PDF page exceeds its available text slots."
    page("Slots").Add Array(value, x, y, width, height, font, size, align, opaque, bold)
End Sub
Private Sub SetMetricFont(ByVal font As String, ByVal size As Double, Optional ByVal bold As Boolean = False)
    Dim driver As String, device As String, key As String, face As String, created As LongPtr, previous As LongPtr
    If metricDC = 0 Then
        driver = "WINSPOOL": device = Application.Printer.DeviceName
        metricDC = PDFCreateDC(StrPtr(driver), StrPtr(device), 0, 0)
        If metricDC = 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Printer font metrics are unavailable. No PDF was generated."
        metricDpiX = PDFDeviceCaps(metricDC, 88): metricDpiY = PDFDeviceCaps(metricDC, 90)
        If metricDpiX <= 0 Or metricDpiY <= 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Printer font metrics have no usable resolution."
    End If
    key = font & "|" & size & "|" & CStr(bold)
    If key = metricKey Then Exit Sub
    created = PDFCreateFont(-CLng(size * metricDpiY / 72#), 0, 0, 0, IIf(bold, 700, 400), 0, 0, 0, 1, 7, 0, 2, 0, StrPtr(font))
    If created = 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Required print font could not be measured: " & font
    previous = PDFSelectObject(metricDC, created)
    If previous = 0 Or previous = -1 Then
        PDFDeleteObject created
        Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Required print font could not be selected."
    End If
    If metricFont = 0 Then metricOldFont = previous Else PDFDeleteObject metricFont
    metricFont = created
    face = String$(128, vbNullChar)
    If PDFTextFace(metricDC, 128, StrPtr(face)) = 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Selected print font could not be verified."
    face = Left$(face, InStr(face, vbNullChar) - 1)
    If StrComp(face, font, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.SetMetricFont", "Required print font is unavailable: " & font & ". Windows substituted " & face & "."
    metricKey = key
End Sub
Public Function PDFOutputTextWidth(ByVal value As String, ByVal font As String, ByVal size As Double, Optional ByVal bold As Boolean = False) As Double
    Dim extent As PDFExtent
    SetMetricFont font, size, bold
    If value = "" Then Exit Function
    If PDFTextExtent(metricDC, StrPtr(value), Len(value), extent) = 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.PDFOutputTextWidth", "Print text could not be measured."
    PDFOutputTextWidth = extent.Width * 72# / metricDpiX
End Function
Private Function PDFLineHeight(ByVal font As String, ByVal size As Double, Optional ByVal bold As Boolean = False, Optional ByVal compact As Boolean = False) As Double
    Dim metrics As PDFFontMetrics
    SetMetricFont font, size, bold
    If PDFTextMetrics(metricDC, metrics) = 0 Then Err.Raise vbObjectError + 9110, "modPDFOutput.PDFLineHeight", "Print line height could not be measured."
    PDFLineHeight = metrics.Height * 72# / metricDpiY
    If Not compact Then PDFLineHeight = PDFLineHeight + metrics.ExternalLeading * 72# / metricDpiY
End Function
Private Sub CloseMetricContext()
    If metricDC <> 0 Then
        If metricOldFont <> 0 Then PDFSelectObject metricDC, metricOldFont
        If metricFont <> 0 Then PDFDeleteObject metricFont
        PDFDeleteDC metricDC
    End If
    metricDC = 0: metricFont = 0: metricOldFont = 0: metricKey = ""
End Sub
'Release printer metrics when the editor closes; a running export keeps its context.
Public Sub PDFFieldFitRelease()
    If Not exporting Then CloseMetricContext
    If gdipToken <> 0 Then PDFGdiplusShutdown gdipToken: gdipToken = 0
End Sub
Private Function LinesFor(ByVal text As String, ByVal width As Double, ByVal font As String, ByVal size As Double, Optional ByVal bold As Boolean = False, Optional ByVal positions As Collection = Nothing) As Collection
    Dim lines As New Collection, parts As Variant, part As Variant, line As String, c As String, i As Long, count As Long, breakAt As Long, partStart As Long, lineStart As Long
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf): parts = Split(text, vbLf)
    partStart = 1
    For Each part In parts
        line = "": i = 1: lineStart = partStart
        Do While i <= Len(part)
            count = 1: c = Mid$(part, i, 1)
            If AscW(c) < 0 Then
                If (AscW(c) And &HFC00) = &HD800 And i < Len(part) Then count = 2: c = Mid$(part, i, 2)
            End If
            If Len(line) > 0 And PDFOutputTextWidth(line & c, font, size, bold) > width - 1 Then
                breakAt = InStrRev(line, " ")
                If breakAt > 1 Then
                    lines.Add Left$(line, breakAt - 1)
                    lineStart = lineStart + breakAt
                    If Not positions Is Nothing Then positions.Add lineStart
                    line = Mid$(line, breakAt + 1)
                Else
                    lines.Add line: lineStart = lineStart + Len(line)
                    If Not positions Is Nothing Then positions.Add lineStart
                    line = ""
                End If
            End If
            line = line & c: i = i + count
        Loop
        lines.Add line
        partStart = partStart + Len(part) + 1
        If Not positions Is Nothing Then positions.Add partStart
    Next
    Set LinesFor = lines
End Function
'Wrap a value into a print box exactly as the PDF does. Shared by FitText and PDFFieldFit.
Private Function BoxLines(ByVal value As String, ByVal box As Variant, ByVal font As String, ByVal size As Double, ByVal bold As Boolean, ByVal compact As Boolean, ByRef step As Double, ByRef capacity As Long, ByRef remaining As String, ByVal positions As Collection) As Collection
    Dim normalized As String
    step = PDFLineHeight(font, size, bold, compact)
    'Retain an exact line despite floating-point division; tolerance is below a twip.
    capacity = Fix((CDbl(box(3)) + 0.0001) / step): If capacity < 0 Then capacity = 0
    remaining = ""
    If value = "" Then Set BoxLines = New Collection: Exit Function
    Set BoxLines = LinesFor(value, CDbl(box(2)), font, size, bold, positions)
    normalized = Replace(Replace(value, vbCrLf, vbLf), vbCr, vbLf)
    If BoxLines.Count > capacity Then
        If capacity = 0 Then remaining = value Else remaining = Mid$(normalized, CLng(positions(capacity)))
    End If
End Function
Private Sub FitText(ByVal page As Object, ByVal value As String, ByVal label As String, ByVal box As Variant, ByVal font As String, ByVal size As Double, ByVal overflow As Collection, Optional ByVal align As Long = 1, Optional ByVal bold As Boolean = False, Optional ByVal spill As Object = Nothing, Optional ByVal compact As Boolean = False)
    Dim lines As Collection, positions As New Collection, i As Long, step As Double, capacity As Long, remaining As String
    If value = "" Then Exit Sub
    Set lines = BoxLines(value, box, font, size, bold, compact, step, capacity, remaining, positions)
    If remaining <> "" And capacity = 1 And spill Is Nothing Then
        remaining = value
        label = label & " (full value; field marked *)"
        AddSlot page, "*", CDbl(box(0)), CDbl(box(1)), CDbl(box(2)), step + 0.5, font, size, align, False, bold
    ElseIf remaining <> "" And spill Is Nothing Then
        label = label & " (continued)"
    End If
    If Not (remaining <> "" And capacity = 1 And spill Is Nothing) Then
        For i = 1 To lines.Count
            If i > capacity Then Exit For
            AddSlot page, CStr(lines(i)), CDbl(box(0)), CDbl(box(1)) + (i - 1) * step, CDbl(box(2)), step + 0.5, font, size, align, False, bold
        Next
    End If
    If remaining <> "" Then
        If spill Is Nothing Then overflow.Add label & vbCrLf & remaining Else spill("Text") = remaining
    End If
End Sub
Private Sub FitGradeExplanation(ByVal page As Object, ByVal value As String, ByVal label As String, ByVal box As Variant, ByVal font As String, ByVal size As Double, ByVal overflow As Collection, ByVal narrowAt As Double, ByVal narrowWidth As Double)
    Dim spill As Object, upper As Variant, lower As Variant, step As Double, height As Double
    If value = "" Then Exit Sub
    step = PDFLineHeight(font, size, False, True): height = Fix((narrowAt - CDbl(box(1)) - 0.5) / step) * step
    Set spill = CreateObject("Scripting.Dictionary")
    upper = Array(box(0), box(1), box(2), height)
    FitText page, value, label, upper, font, size, overflow, 1, False, spill, True
    If spill.Exists("Text") Then
        lower = Array(box(0), CDbl(box(1)) + height, narrowWidth, CDbl(box(3)) - height)
        FitText page, CStr(spill("Text")), label, lower, font, size, overflow, 1, False, Nothing, True
    End If
End Sub
'Measure a report field against its printed box for the editor line meter, using the
'same print rows, wrap and capacity as PlanReport. Never raises: a missing printer or
'font gives Status "Unavailable" with Reason; a field without a plain text box gives "None".
Public Function PDFFieldFit(ByVal kind As String, ByVal key As String, ByVal text As String) As Object
    Dim result As Object, row As Variant, parts As Variant, number As Long, found As Boolean
    Dim box As Variant, font As String, size As Double, bold As Boolean, step As Double, capacity As Long
    Dim lines As Collection, positions As New Collection, remaining As String, normalized As String, count As Long, last As String
    On Error GoTo Unavailable
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Status", "None": result.Add "Lines", 0&: result.Add "Capacity", 0&
    result.Add "Overflow", 0&: result.Add "OneWordLast", False: result.Add "Reason", ""
    Set PDFFieldFit = result
    For number = 1 To 2
        For Each row In PDFPrintRows(kind, number)
            parts = Split(CStr(row), "|")
            If parts(0) = key And parts(9) <> "1" And Val(parts(13)) <= 0 Then found = True: Exit For
        Next
        If found Then Exit For
    Next
    If Not found Then Exit Function
    box = Array(Val(parts(3)), Val(parts(4)), Val(parts(5)), Val(parts(6)))
    font = parts(7): size = Val(parts(8)): bold = (parts(11) = "1")
    Set lines = BoxLines(text, box, font, size, bold, False, step, capacity, remaining, positions)
    result("Capacity") = capacity: count = lines.Count
    normalized = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    If Right$(normalized, 1) = vbLf Then count = count - 1
    result("Lines") = count
    If remaining <> "" Then
        'A one-line box moves the whole value to the added page (FitText's "*" case).
        If capacity = 1 Then result("Overflow") = count Else result("Overflow") = count - capacity
        If result("Overflow") < 1 Then result("Overflow") = 1&
    End If
    If result("Overflow") = 0 And count >= 2 Then
        last = Trim$(CStr(lines(count)))
        result("OneWordLast") = (last <> "" And InStr(last, " ") = 0)
    End If
    result("Status") = "Measured"
    Exit Function
Unavailable:
    If result Is Nothing Then Exit Function
    result("Status") = "Unavailable": result("Reason") = Err.Description
    result("Lines") = 0&: result("Capacity") = 0&: result("Overflow") = 0&: result("OneWordLast") = False
End Function
'Lines a report field takes in NAVFIT98A 2.2.0.332's print at default comments pitch. Swift
'never stores CommentsPitch, so its exports print at the default. Fonts and widths are the
'PrintModule rectangles; capacities are NAVFIT98A's own editor limits, whose warnings say
'longer text will be truncated. See planning/NAVFIT_PRINT_ENGINE_FINDINGS.md. Never raises.
Public Function NAVFITFieldFit(ByVal kind As String, ByVal key As String, ByVal text As String) As Object
    Dim result As Object, size As Single, width As Single, capacity As Long, prefix As String, count As Long, shorter As String, cut As Long
    On Error GoTo Unavailable
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Status", "None": result.Add "Lines", 0&: result.Add "Capacity", 0&
    result.Add "Overflow", 0&: result.Add "OneWordLast", False: result.Add "Reason", ""
    Set NAVFITFieldFit = result
    If kind <> "EVAL" And kind <> "FITREP" And kind <> "CHIEFEVAL" Then Exit Function
    Select Case key
        Case "OfficialComments"
            If kind = "CHIEFEVAL" Then size = 10: width = 760: capacity = 6 Else size = 9: width = 715: capacity = 18
        Case "CommandEmployment": size = 9.65: width = 755: capacity = 3
        Case "DutyDescription": size = 9.65: width = 755: capacity = 4: prefix = Space$(21)
        Case Else
            If kind <> "CHIEFEVAL" Or Not key Like "Trait3#Comment#" Then Exit Function
            size = 9.65: width = 450: capacity = 2
    End Select
    result("Capacity") = capacity
    If text <> "" Then
        count = NAVFITLines(prefix & text, size, width)
        result("Lines") = count
        If count > capacity Then result("Overflow") = count - capacity
        'The last line holds one word when dropping that word removes the line. Text that
        'ends with a line break ends on a blank line.
        If count >= 2 And count <= capacity And Right$(text, 1) <> vbLf And Right$(text, 1) <> vbCr Then
            shorter = RTrim$(Replace(Replace(text, vbCr, " "), vbLf, " "))
            cut = InStrRev(shorter, " ")
            If cut > 0 Then
                shorter = RTrim$(Left$(text, cut - 1))
                If Right$(shorter, 1) = vbLf Or Right$(shorter, 1) = vbCr Then
                    result("OneWordLast") = True
                Else
                    result("OneWordLast") = (NAVFITLines(prefix & shorter, size, width) = count - 1)
                End If
            End If
        End If
    End If
    result("Status") = "Measured"
    Exit Function
Unavailable:
    If result Is Nothing Then Exit Function
    result("Status") = "Unavailable": result("Reason") = Err.Description
    result("Lines") = 0&: result("Capacity") = 0&: result("Overflow") = 0&: result("OneWordLast") = False
End Function
'GDI+ line count for Courier New text as DrawString(text, font, brush, RectangleF) lays it out:
'no StringFormat, unlimited height, page unit Display (1/100 inch) on the printer context.
Private Function NAVFITLines(ByVal text As String, ByVal size As Single, ByVal width As Single) As Long
    Dim options As PDFGdiplusInput, graphics As LongPtr, family As LongPtr, font As LongPtr, layout As PDFRectF, bounds As PDFRectF
    Dim fitted As Long, lines As Long, status As Long, face As String, failure As Long, message As String
    On Error GoTo Failed
    SetMetricFont "Courier New", 10
    If gdipToken = 0 Then
        options.Version = 1
        If PDFGdiplusStartup(gdipToken, options, 0) <> 0 Then gdipToken = 0: Err.Raise vbObjectError + 9112, "modPDFOutput.NAVFITLines", "Windows GDI+ is unavailable."
    End If
    face = "Courier New"
    status = PDFGdipCreateFromHDC(metricDC, graphics)
    If status = 0 Then status = PDFGdipSetPageUnit(graphics, 1)
    If status = 0 Then status = PDFGdipCreateFontFamily(StrPtr(face), 0, family)
    If status = 0 Then status = PDFGdipCreateFont(family, size, 0, 3, font)
    layout.Width = width: layout.Height = 1000000
    If status = 0 Then status = PDFGdipMeasureString(graphics, StrPtr(text), Len(text), font, layout, 0, bounds, fitted, lines)
    If status <> 0 Then Err.Raise vbObjectError + 9112, "modPDFOutput.NAVFITLines", "NAVFIT98A line measurement failed (GDI+ status " & status & ")."
    NAVFITLines = lines
Finished:
    If font <> 0 Then PDFGdipDeleteFont font: font = 0
    If family <> 0 Then PDFGdipDeleteFontFamily family: family = 0
    If graphics <> 0 Then PDFGdipDeleteGraphics graphics: graphics = 0
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, "modPDFOutput.NAVFITLines", message
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    Resume Finished
End Function
Private Function TraitAverage(ByVal fields As Object) As String
    Dim key As Variant, mark As String, total As Double, count As Long
    If PDFOutputText(fields, "ObservationStatus") = "NOB" Then Exit Function
    For Each key In TraitKeys()
        mark = PDFOutputText(fields, CStr(key))
        If mark Like "[1-5].0" Then total = total + Val(mark): count = count + 1
    Next
    If count > 0 Then TraitAverage = Format$(total / count, "0.00")
End Function
Private Function GroupAverage(ByVal rows As Collection) As String
    Dim row As Object, fields As Object, key As Variant, mark As String, total As Double, count As Long
    For Each row In rows
        Set fields = row("Fields")
        If PDFOutputText(fields, "ObservationStatus") <> "NOB" Then
            For Each key In TraitKeys()
                mark = PDFOutputText(fields, CStr(key))
                If mark Like "[1-5].0" Then total = total + Val(mark): count = count + 1
            Next
        End If
    Next
    If count > 0 Then GroupAverage = Format$(total / count, "0.00")
End Function
Private Function RecommendationCount(ByVal rows As Collection, ByVal index As Long) As Long
    Dim row As Object, names As Variant
    names = Array("Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
    For Each row In rows
        If PDFOutputText(row("Fields"), "ObservationStatus") <> "NOB" And PDFOutputText(row("Fields"), "PromotionRecommendation") = names(index) Then RecommendationCount = RecommendationCount + 1
    Next
End Function
Private Function SeniorCity(ByVal fields As Object) As String
    SeniorCity = Trim$(PDFOutputText(fields, "ReportingSeniorCity") & " " & PDFOutputText(fields, "ReportingSeniorState") & " " & PDFOutputText(fields, "ReportingSeniorZip"))
End Function
Private Function PrintNavyDate(ByVal value As String, Optional ByVal readiness As Boolean = False) As String
    Dim months As Variant, month As Long
    PrintNavyDate = value
    If Not IsValidIsoDate(value) Then Exit Function
    month = CLng(Mid$(value, 6, 2))
    If readiness Then
        PrintNavyDate = Mid$(value, 6, 2) & "/" & Mid$(value, 3, 2)
    Else
        months = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
        PrintNavyDate = Mid$(value, 3, 2) & months(month - 1) & Right$(value, 2)
    End If
End Function
Private Function FieldValue(ByVal row As Object, ByVal key As String, ByVal rule As String, ByVal selector As String, ByVal rows As Collection) As String
    Dim fields As Object, value As String, names As Variant, index As Long, parts As Variant, text As String, line As Variant, found As Long, block As Long
    Set fields = row("Fields")
    Select Case key
        Case "@ITA": FieldValue = TraitAverage(fields): Exit Function
        Case "@GroupAverage": If job("StatisticsGroupID") > 0 Then FieldValue = GroupAverage(job("StatisticsRows"))
            Exit Function
        Case "@SeniorCity": FieldValue = SeniorCity(fields): Exit Function
        Case "@SeniorAddress"
            For Each line In Array(PDFOutputText(fields, "ReportingSeniorAddress1"), PDFOutputText(fields, "ReportingSeniorAddress2"), SeniorCity(fields), PDFOutputText(fields, "ReportingSeniorPhone"), PDFOutputText(fields, "ReportingSeniorDSN"))
                If Len(line) > 0 Then
                    If text <> "" Then text = text & vbCrLf
                    text = text & line
                End If
            Next
            FieldValue = text: Exit Function
    End Select
    If Left$(key, 6) = "@Count" Then
        If job("StatisticsGroupID") > 0 Then FieldValue = CStr(RecommendationCount(job("StatisticsRows"), CLng(Right$(key, 1))))
        Exit Function
    End If
    If Left$(key, 6) = "@Grade" Then
        block = CLng(Mid$(key, 7, 2)): value = PDFOutputText(fields, "Trait" & block)
        If value = "" Or value = "NOB" Then Exit Function
        If selector <> "" And InStr(1, "," & selector & ",", "," & value & ",", vbBinaryCompare) = 0 Then Exit Function
        FieldValue = CatalogueTraitDescription(row("Kind"), block, value)
        Exit Function
    End If
    value = PDFOutputText(fields, key)
    If key = "OfficialName" And value = "" Then value = PDFOutputText(fields, "MemberName")
    If key = "MemberSSN" And value = "" Then value = PDFOutputText(fields, "MemberDoDID")
    If row("Kind") = "CHIEFEVAL" Then
        Select Case key
            Case "PeriodStart", "PeriodEnd", "DateReported", "CounselingDate": value = PrintNavyDate(value)
            Case "ReadinessDate": value = PrintNavyDate(value, True)
            Case "MemberSSN", "ReportingSeniorSSN"
                If Len(value) = 9 And value Like "#########" Then value = Left$(value, 3) & "-" & Mid$(value, 4, 2) & "-" & Right$(value, 4)
        End Select
    End If
    Select Case rule
        Case "flag": If value = "Yes" Then FieldValue = "X"
        Case "observation": If value = "NOB" Then FieldValue = "X"
        Case "trait"
            If (value = "NOB" And LCase$(selector) = "nob") Or (value Like "[1-5].0" And selector = Left$(value, 1)) Then FieldValue = "X"
        Case "chief-trait"
            names = Array("Below Standards", "Progressing", "Meets Standards", "Above Standards", "Greatly Exceeds Standards")
            If value Like "[1-5].0" Then
                FieldValue = value & " " & names(CLng(Left$(value, 1)) - 1)
            Else
                FieldValue = value
            End If
        Case "chief-mark"
            If value = "NOB" Then FieldValue = "0.0" Else FieldValue = value
        Case "recommendation"
            names = Array("NOB", "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
            parts = Split("NOB|sp|pg|pm|mp|ep", "|")
            For index = 0 To 5
                If value = names(index) And LCase$(selector) = LCase$(parts(index)) Then FieldValue = "X"
            Next
        Case Else
            FieldValue = value
    End Select
End Function
Private Sub AddContinuations(ByVal heading As String, ByVal overflow As Collection)
    Dim item As Variant, lines As Collection, line As Variant, page As Object, y As Double
    For Each item In overflow
        Set lines = LinesFor(CStr(item), 540, "Courier New", 10)
        For Each line In lines
            If page Is Nothing Then
                Set page = NewPage(""): y = 66
                AddSlot page, heading & " - additional text", 36, 24, 540, 28, "Arial", 11
            End If
            AddSlot page, CStr(line), 36, y, 540, 15, "Courier New", 10
            y = y + 14
            If y > 736 Then Set page = Nothing
        Next
        y = y + 14: If y > 736 Then Set page = Nothing
    Next
End Sub
Private Function PrintFieldLabel(ByVal key As String) As String
    Dim names As Variant
    If Left$(key, 6) = "@Count" Then
        names = Array("Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
        PrintFieldLabel = "Summary promotion distribution - " & names(CLng(Right$(key, 1))): Exit Function
    End If
    If Left$(key, 6) = "@Grade" Then
        PrintFieldLabel = "Trait " & Mid$(key, 7, 2) & " performance grade explanation": Exit Function
    End If
    Select Case key
        Case "@SeniorCity", "@SeniorAddress": PrintFieldLabel = "Reporting senior address and contact details"
        Case "@ITA": PrintFieldLabel = "Individual trait average"
        Case "@GroupAverage": PrintFieldLabel = "Summary group average"
        Case Else: PrintFieldLabel = UXFieldLabel(key)
    End Select
End Function
Private Sub PlanReport(ByVal row As Object, ByVal rows As Collection)
    Dim page As Object, number As Long, map As Variant, parts As Variant, value As String, box As Variant, extra As New Collection, covered As Object, key As Variant, unrepresented As Object, raw As String, valid As Boolean, align As Long, lineHeight As Double
    Set covered = CreateObject("Scripting.Dictionary")
    Set unrepresented = CreateObject("Scripting.Dictionary")
    If row("Kind") <> "EVAL" And row("Kind") <> "CHIEFEVAL" And row("Kind") <> "FITREP" Then Err.Raise vbObjectError + 9106, "modPDFOutput.PlanReport", "No print template exists for report type " & row("Kind") & "."
    For number = 1 To 2
        Set page = NewPage(row("Kind") & "_" & number)
        For Each map In PDFPrintRows(CStr(row("Kind")), number)
            parts = Split(CStr(map), "|")
            value = FieldValue(row, CStr(parts(0)), CStr(parts(1)), CStr(parts(2)), rows)
            If Left$(parts(0), 1) <> "@" Then covered(CStr(parts(0))) = True
            raw = PDFOutputText(row("Fields"), CStr(parts(0))): valid = True
            Select Case parts(1)
                Case "flag": valid = (raw = "" Or raw = "Yes" Or raw = "No")
                Case "observation": valid = (raw = "" Or raw = "Observed" Or raw = "NOB")
                Case "trait", "chief-trait", "chief-mark": valid = (raw = "" Or raw = "NOB" Or raw Like "[1-5].0")
                Case "recommendation": valid = (raw = "" Or raw = "NOB" Or raw = "Significant Problems" Or raw = "Progressing" Or raw = "Promotable" Or raw = "Must Promote" Or raw = "Early Promote")
            End Select
            If Not valid And Not unrepresented.Exists(CStr(parts(0))) Then
                extra.Add PrintFieldLabel(CStr(parts(0))) & " (unrecognized saved/draft value)" & vbCrLf & raw
                unrepresented(CStr(parts(0))) = True
            End If
            Select Case parts(0)
                Case "@SeniorCity": covered("ReportingSeniorCity") = True: covered("ReportingSeniorState") = True: covered("ReportingSeniorZip") = True
                Case "@SeniorAddress"
                    For Each key In Array("ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN"): covered(CStr(key)) = True: Next
            End Select
            box = Array(Val(parts(3)), Val(parts(4)), Val(parts(5)), Val(parts(6)))
            align = 1
            If CStr(parts(0)) = "@ITA" Or CStr(parts(0)) = "@GroupAverage" Or Left$(parts(0), 6) = "@Count" Then
                align = 2: lineHeight = Int(Val(parts(8)) + 0.5) * 1.2
                If CDbl(box(3)) > lineHeight Then
                    box(1) = CDbl(box(1)) + (CDbl(box(3)) - lineHeight) / 2: box(3) = lineHeight
                End If
            End If
            If parts(9) = "1" Then
                AddSlot page, value, CDbl(box(0)), CDbl(box(1)), CDbl(box(2)), CDbl(box(3)), "Arial", IIf(row("Kind") = "CHIEFEVAL", 8, 10), 2
            ElseIf Val(parts(13)) > 0 Then
                FitGradeExplanation page, value, CStr(row("Caption")) & " / " & PrintFieldLabel(CStr(parts(0))), box, CStr(parts(7)), Val(parts(8)), extra, Val(parts(12)), Val(parts(13))
            Else
                If parts(1) = "chief-mark" Then align = 2
                FitText page, value, CStr(row("Caption")) & " / " & PrintFieldLabel(CStr(parts(0))), box, CStr(parts(7)), Val(parts(8)), extra, align, (parts(11) = "1")
            End If
        Next
    Next
    'Private notes and name components are deliberately not repeated as addenda.
    For Each key In DraftFieldKeys()
        If Not covered.Exists(CStr(key)) And PDFOutputText(row("Fields"), CStr(key)) <> "" Then
            If PDFPrintAdditionalField(CStr(key)) Then extra.Add PrintFieldLabel(CStr(key)) & vbCrLf & PDFOutputText(row("Fields"), CStr(key))
        End If
    Next
    AddContinuations CStr(row("Caption")), extra
End Sub
Private Function CommonField(ByVal rows As Collection, ByVal key As String) As String
    Dim row As Object, value As String, first As Boolean
    first = True
    For Each row In rows
        value = PDFOutputText(row("Fields"), key)
        If first Then
            CommonField = value: first = False
        ElseIf CommonField <> value Then
            CommonField = "MIXED": Exit Function
        End If
    Next
End Function
Private Function SummaryFlags(ByVal rows As Collection, ByVal keys As Variant, ByVal labels As Variant) As String
    Dim i As Long, value As String
    For i = LBound(keys) To UBound(keys)
        value = CommonField(rows, CStr(keys(i)))
        If value = "MIXED" Then SummaryFlags = "MIXED": Exit Function
        If value = "Yes" Then
            If SummaryFlags <> "" Then SummaryFlags = SummaryFlags & "/"
            SummaryFlags = SummaryFlags & labels(i)
        End If
    Next
End Function
Private Sub PlanSummary(ByVal rows As Collection)
    Dim page As Object, template As Long, bounds As Variant, i As Long, col As Long, index As Long, n As Long, x As Double, row As Object, extra As New Collection, j As Long, key As Variant, values As Variant, boxes As Variant, total As Long, count As Long
    template = 1: index = 1
    Do
        Set page = NewPage("SUMMARY_" & template): bounds = PDFSummaryRows(template): n = UBound(bounds)
        FitText page, CStr(job("Title")), "Summary filename / group", Array(365, 30, 201, 14), "Courier New", 9, extra
        values = Array(CommonField(rows, "ReportingSeniorName"), CommonField(rows, "ReportingSeniorGrade"), CommonField(rows, "ReportingSeniorSSN"), CommonField(rows, "ReportingSeniorUIC"), job("Paygrade"), CommonField(rows, "PromotionStatus"), CommonField(rows, "Designator"), SummaryFlags(rows, Array("ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW"), Array("ACT", "TAR", "INACT", "AT")), CommonField(rows, "BilletSubcategory"), SummaryFlags(rows, Array("ReportRegular", "ReportConcurrent", "ReportOperationalCommander"), Array("REGULAR", "CONCURRENT", "OPS CDR")), CommonField(rows, "UnitUIC"), PrintNavyDate(CommonField(rows, "PeriodEnd")))
        boxes = Array(Array(20, 54, 159, 12), Array(185, 54, 173, 12), Array(366, 54, 91, 12), Array(462, 54, 104, 12), Array(20, 77, 157, 12), Array(186, 77, 174, 12), Array(365, 77, 91, 12), Array(463, 77, 104, 12), Array(20, 102, 157, 12), Array(186, 102, 174, 12), Array(365, 102, 91, 12), Array(463, 102, 104, 12))
        For j = 0 To UBound(values): FitText page, CStr(values(j)), "Summary header " & (j + 1), boxes(j), "Times New Roman", 9, extra: Next
        For j = 0 To 4: AddSlot page, CStr(RecommendationCount(rows, j)), 165 + j * 80, 140, 72, 22, "Times New Roman", 11, 2: Next
        For col = 0 To 1
            For i = 0 To n - 1
                If index <= rows.Count Then
                    Set row = rows(index): x = IIf(col = 0, 18, 298)
                    FitText page, CStr(row("Last")), CStr(row("Caption")) & " / summary name", Array(x, CDbl(bounds(i)) + 0.5, IIf(col = 0, 117, 119), CDbl(bounds(i + 1)) - CDbl(bounds(i)) - 0.7), "Times New Roman", 8, extra
                    AddSlot page, Right$(PDFOutputText(row("Fields"), "MemberSSN"), 4), IIf(col = 0, 139, 421), CDbl(bounds(i)) + 0.5, 68, 10, "Times New Roman", 8
                    AddSlot page, TraitAverage(row("Fields")), IIf(col = 0, 215, 492), CDbl(bounds(i)) + 0.5, 70, 10, "Times New Roman", 8
                    index = index + 1
                End If
            Next
        Next
        If template = 1 Then
            AddSlot page, CStr(rows.Count), 492, 482, 75, 14, "Times New Roman", 11, 2
            AddSlot page, GroupAverage(rows), 492, 506, 75, 14, "Times New Roman", 11, 2
            If index <= rows.Count Then AddSlot page, "X", 549, 529, 20, 14, "Arial", 10, 2
        End If
        If template = 1 Then AddSlot page, " ", 503, 574, 69, 18, "Times New Roman", 7, 1, True
        If template = 1 Then PlanSummaryContact page, rows, extra
        template = 2
    Loop While index <= rows.Count
    AddContinuations CStr(job("Title")) & " summary", extra
    total = printPages.Count
    For i = 1 To total
        AddSlot printPages(i), "Page " & i & " of " & total, 519, 736, 53, 14, "Times New Roman", 8, 3, True
    Next
End Sub
'The letter's lower section: submitted on time, ISIC address and UIC, the reporting
'senior's address, e-mail and phone, and the date the letter was created.
Private Sub PlanSummaryContact(ByVal page As Object, ByVal rows As Collection, ByVal extra As Collection)
    Dim common As Object, details As Object, key As Variant
    Set common = CreateObject("Scripting.Dictionary")
    For Each key In Array("ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone")
        common.Add CStr(key), CommonField(rows, CStr(key))
    Next
    Set details = SummaryLetterDetails(job("GroupID"), job("ISICID"), CStr(job("OnTime")), common)
    If details("OnTime") = "Yes" Then AddSlot page, "X", 180, 453, 21, 14, "Arial", 10, 2
    If details("OnTime") = "No" Then AddSlot page, "X", 227, 453, 21, 14, "Arial", 10, 2
    FitText page, CStr(details("ISICAddress")), "Summary ISIC address", Array(22, 592, 210, 44), "Times New Roman", 8, extra
    AddSlot page, CStr(details("ISICUIC")), 200, 576, 34, 10, "Times New Roman", 8
    FitText page, CStr(details("SeniorAddress")), "Summary reporting senior address", Array(22, 663, 210, 44), "Times New Roman", 8, extra
    AddSlot page, CStr(details("Email")), 47, 720.5, 190, 9, "Times New Roman", 8
    AddSlot page, CStr(details("Phone")), 47, 729.5, 190, 9, "Times New Roman", 8
    AddSlot page, PrintNavyDate(Format$(job("Captured"), "yyyy-mm-dd")), 47, 738.5, 190, 9, "Times New Roman", 8
End Sub
Private Sub WritePageCache()
    Dim db As DAO.Database, r As DAO.Recordset, sql As String, number As Long, page As Object
    sql = "CREATE TABLE PDFPage (PageNo LONG PRIMARY KEY)"
    Set db = DBEngine.CreateDatabase(cachePath, dbLangGeneral, dbVersion120): db.Execute sql, dbFailOnError
    Set r = db.OpenRecordset("PDFPage", dbOpenDynaset)
    For Each page In printPages
        number = number + 1: r.AddNew: r!PageNo = number
        r.Update
    Next
    r.Close: db.Close: Set r = Nothing: Set db = Nothing
End Sub
Private Sub EmitPDF(ByVal destination As String)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(destination) Or fs.FolderExists(destination) Then Err.Raise vbObjectError + 9107, "modPDFOutput.EmitPDF", "PDF destination already exists. Existing output was preserved."
    Call WritePageCache: outputFailure = ""
    DoCmd.OpenReport "rptPDFOutput", acViewPreview, , , acHidden
    DoCmd.OutputTo acOutputReport, "rptPDFOutput", acFormatPDF, destination, False, , , acExportQualityPrint
    ClosePrintReport
    If outputFailure <> "" Then Err.Raise vbObjectError + 9108, "modPDFOutput.EmitPDF", outputFailure
    If Not fs.FileExists(destination) Then Err.Raise vbObjectError + 9109, "modPDFOutput.EmitPDF", "Access did not produce a PDF."
    If fs.GetFile(destination).Size < 100 Then Err.Raise vbObjectError + 9110, "modPDFOutput.EmitPDF", "Access produced an empty PDF."
    fs.DeleteFile cachePath: cachePath = ""
End Sub
Private Function JSONString(ByVal text As String) As String
    text = Replace(Replace(Replace(Replace(Replace(text, "\", "\\"), """", "\"""), vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
    JSONString = """" & text & """"
End Function
Private Sub WriteManifest(ByVal folder As String, ByVal rows As Collection, ByVal files As Collection)
    Dim text As String, row As Object, file As Variant, stream As Object, i As Long
    text = "{""format"":1,""captured"":" & JSONString(Format$(job("Captured"), "yyyy-mm-dd hh:nn:ss")) & ",""application"":" & JSONString(APP_VERSION) & ",""layout"":" & JSONString(PDFPrintCatalogueHash()) & ",""reports"":["
    For Each row In rows
        If i > 0 Then text = text & ","
        text = text & "{""guid"":" & JSONString(row("Guid")) & ",""revision"":" & row("Revision") & ",""source"":" & JSONString(row("Source")) & "}": i = i + 1
    Next
    text = text & "],""files"":[": i = 0
    For Each file In files
        If i > 0 Then text = text & ","
        text = text & "{""name"":" & JSONString(CStr(file)) & ",""sha256"":" & JSONString(SourceFileHash(folder & "\" & file)) & "}": i = i + 1
    Next
    text = text & "]}"
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text: stream.SaveToFile folder & "\manifest.json", 1: stream.Close
End Sub
Public Function PDFOutputGenerate(ByVal root As String, ByVal selected As Object, ByVal mode As String, ByVal includeSummary As Boolean) As String
    Dim rows As New Collection, row As Object, fs As Object, groupFolder As String, finalFolder As String, stageFolder As String, operation As Long, file As String, files As New Collection, stage As String, failure As Long, message As String, cleanup As String
    On Error GoTo Failed
    If exporting Then Err.Raise vbObjectError + 9111, "modPDFOutput.PDFOutputGenerate", "A PDF export is already running."
    If job Is Nothing Then Err.Raise vbObjectError + 9100, "modPDFOutput.PDFOutputGenerate", "Refresh the PDF preview."
    PDFOutputEndPreview
    If mode <> "Combined packet" And mode <> "Individual PDFs" And mode <> "Both" Then Err.Raise 5, "modPDFOutput.PDFOutputGenerate", "Choose an output format."
    For Each row In job("Reports")
        If selected.Exists(CStr(row("ID"))) Then rows.Add row
    Next
    If includeSummary And job("GroupID") = 0 Then includeSummary = False
    If rows.Count = 0 And Not includeSummary Then Err.Raise 5, "modPDFOutput.PDFOutputGenerate", "Select a report or the summary sheet."
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(root) Then Err.Raise 76, "modPDFOutput.PDFOutputGenerate", "Choose an available export folder."
    exporting = True: stage = "Create output run": job("Folder") = ""
    groupFolder = fs.BuildPath(root, PDFOutputSafePart(job("Title")))
    If Not fs.FolderExists(groupFolder) Then fs.CreateFolder groupFolder
    finalFolder = fs.BuildPath(groupFolder, Format$(Now(), "yyyy-mm-dd_hhnnss") & "_" & ShortIdentity(NewIdentity()))
    stageFolder = finalFolder & ".partial": fs.CreateFolder stageFolder
    operation = BeginOperation("Create print-ready PDFs", "Captured PDF preview " & job("ID") & "; reports " & rows.Count)
    stage = "Plan print pages"
    If mode = "Combined packet" Or mode = "Both" Then
        Set printPages = New Collection
        If includeSummary Then PlanSummary rows
        For Each row In rows: PlanReport row, rows: Next
        Call ClosePrintReport: stage = "Write combined packet"
        cachePath = stageFolder & "\pages.accdb": file = PDFOutputSafePart(job("Title")) & "__Packet.pdf": EmitPDF stageFolder & "\" & file: files.Add file
    End If
    If mode = "Individual PDFs" Or mode = "Both" Then
        If includeSummary Then
            Set printPages = New Collection: PlanSummary rows: ClosePrintReport
            stage = "Write summary sheet": cachePath = stageFolder & "\pages.accdb": file = PDFOutputSafePart(job("Title")) & "__Summary_Letter.pdf": EmitPDF stageFolder & "\" & file: files.Add file
        End If
        For Each row In rows
            Set printPages = New Collection: PlanReport row, rows: ClosePrintReport
            stage = "Write individual report": cachePath = stageFolder & "\pages.accdb": file = PDFOutputFileName(job("Title"), row): EmitPDF stageFolder & "\" & file: files.Add file
        Next
    End If
    Call ClosePrintReport: stage = "Verify output inventory"
    If files.Count = 0 Then Err.Raise 5, "modPDFOutput.PDFOutputGenerate", "No PDF documents were selected."
    WriteManifest stageFolder, rows, files
    stage = "Publish completed run": fs.MoveFolder stageFolder, finalFolder
    job("Folder") = finalFolder: Set job("Files") = files
    FinishOperation operation, "Completed", files.Count & " print-ready PDFs; " & rows.Count & " captured reports; output " & finalFolder
    exporting = False: PDFOutputGenerate = finalFolder: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    ClosePrintReport
    If Err.Number <> 0 Then cleanup = vbCrLf & "Print-report cleanup: " & Err.Number & " " & Err.Description
    Err.Clear
    If operation > 0 Then FinishOperation operation, "Failed", message
    If Err.Number <> 0 Then cleanup = cleanup & vbCrLf & "Operation-journal cleanup: " & Err.Number & " " & Err.Description
    On Error GoTo 0
    exporting = False
    If stage = "Publish completed run" Then
        If job("Folder") <> "" Then message = message & vbCrLf & "PDF files were published at " & job("Folder") & "; recording the outcome failed."
    ElseIf stageFolder <> "" Then
        message = message & vbCrLf & "Incomplete run retained at " & stageFolder
    End If
    Err.Raise failure, "PDFOutputGenerate", message & cleanup
End Function
