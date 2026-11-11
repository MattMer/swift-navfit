Attribute VB_Name = "PDFOutputHarness"
Option Explicit
Private n As Long

Public Function TestPDFMetrics() As String
    Dim wide As Double, narrow As Double, courier As Double
    On Error GoTo Failed
    wide = PDFOutputTextWidth("WWWWWWWWWW", "Times New Roman", 10)
    narrow = PDFOutputTextWidth("iiiiiiiiii", "Times New Roman", 10)
    courier = PDFOutputTextWidth("WWWWWWWWWW", "Courier New", 10)
    If wide <= narrow * 2 Or courier < 50 Or courier > 70 Then Err.Raise 5, "PDFOutputHarness.TestPDFMetrics", "Native report font metrics are not usable in points."
    PDFOutputEndPreview
    TestPDFMetrics = "PASS: native printer font metrics; Times wide=" & wide & "; narrow=" & narrow & "; Courier=" & courier
    Exit Function
Failed:
    TestPDFMetrics = "FAIL: native report metrics: " & Err.Number & " " & Err.Description
    On Error Resume Next
    PDFOutputEndPreview
End Function
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 9190, , message
 n = n + 1
End Sub
Public Function TestChiefRenderingProof(ByVal root As String) As String
 Dim id As Long, block As Long, index As Long, current As Object, rows As Collection, fields As Object, selected As Object, changes As Object, folder As String
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC CHIEF PRINT PROOF", "Synthetic tester"
 id = CommitDraft(0, "CHIEFEVAL", "Synthetic Chief proof", "SYNTHETIC, CHIEF", #1/1/2026#, #9/30/2026#, "", "Synthetic rendering proof", "Synthetic reporting senior narrative.", , Array("1.0", "2.0", "3.0", "4.0", "5.0", "1.0", "2.0"), Array("Observed", "Promotable"))
 Set changes = CreateObject("Scripting.Dictionary")
 changes("OfficialName") = "SYNTHETIC, CHIEF": changes("RankRate") = "YNC": changes("MemberSSN") = "000000000"
 changes("CommandEmployment") = "Synthetic command employment paragraph for native text measurement and placement."
 changes("DutyDescription") = "Synthetic primary duties, qualifications, collateral duties and performance evidence."
 For block = 33 To 39
  For index = 1 To 3
   changes("Trait" & block & "Comment" & index) = "SYNTHETIC authored comment " & block & "." & index & ": distinct member performance evidence."
  Next
 Next
 Call CommitSavedFields(id, CurrentRevision(id), changes, "Synthetic authored Chief comments")
 PDFOutputCapture 0, id, False
 Set selected = CreateObject("Scripting.Dictionary"): selected.Add CStr(id), True
 folder = PDFOutputGenerate(root, selected, "Individual PDFs", False)
 TestChiefRenderingProof = "PASS: Chief rendering proof: " & folder: Exit Function
Failed:
 TestChiefRenderingProof = "FAIL: Chief rendering proof: " & Err.Number & " " & Err.Description
End Function

Public Function TestChiefComments(ByVal root As String, ByVal template As String) As String
 Dim id As Long, block As Long, entry As Long, revision As Long, authored As Long, recovery As Long, start As Long
 Dim changes As Object, expected As Object, fields As Object, selected As Object, rows As Collection, current As Object, surface As Access.Form
 Dim db As DAO.Database, native As DAO.Recordset, prefixes As Variant, key As Variant, column As String, file As String, returned As String, text As String, folder As String
 On Error GoTo Failed
 start = n
 If DCount("*", "UserProfile") > 0 Then BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
 ConfigureWorkspace "SYNTHETIC CHIEF COMMENTS", "Synthetic tester"
 id = CommitDraft(0, "CHIEFEVAL", "Synthetic independent comments", "CHIEF, SYNTHETIC", Null, Null, "", "Synthetic Chief fields")
 Set changes = CreateObject("Scripting.Dictionary"): Set expected = CreateObject("Scripting.Dictionary")
 For Each key In DraftFieldKeys()
  If UXFieldFlag(CStr(key)) Then changes(CStr(key)) = "No"
 Next
 changes("ObservationStatus") = "Observed"
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   changes(key) = "AUTHORED " & block & "." & entry & " O'Brien & evidence": expected(key) = changes(key)
  Next
 Next
 Call CommitSavedFields(id, CurrentRevision(id), changes, "21 independent authored comments")
 authored = CurrentRevision(id)
 Set changes = CreateObject("Scripting.Dictionary"): changes("Trait33") = "5.0"
 Call CommitSavedFields(id, CurrentRevision(id), changes, "Rating change cannot author comments")
 Set fields = SavedFields(id)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check SameFieldValue(fields(key), expected(key)), "Rating change overwrote " & key
   Check MajorRevisionValue(CurrentRevision(id), key) = expected(key), "Snapshot omitted " & key
   Check RevisionFieldStatus(authored, CurrentRevision(id), CStr(key)) = "Unchanged", "Rating-only edit changed comment history: " & key
  Next
 Next
 prefixes = Array("LEAD", "TAC", "PROF", "MIS", "EO", "TEAM", "MIL")
 file = root & "\chief-comments.accdb": revision = CurrentRevision(id)
 Check NAVFITExportReport(id, file, template) = file, "Chief native export failed"
 Set db = DBEngine.OpenDatabase(file, False, True): Set native = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry: column = prefixes(block - 33) & "DN" & entry
   Check native.Fields(column) = expected(key), "Native comment order/mapping differs: " & key
  Next
 Next
 native.Close: db.Close
 Check CurrentRevision(id) = revision, "Native export created a revision"
 returned = root & "\chief-comments-returned.accdb": FileCopy file, returned
 Set db = DBEngine.OpenDatabase(returned)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry: column = prefixes(block - 33) & "DN" & entry
   expected(key) = "RETURNED " & block & "." & entry
   db.Execute "UPDATE Reports SET [" & column & "]=" & QuoteSQL(CStr(expected(key))) & " WHERE ReportType Is Not Null", dbFailOnError
  Next
 Next
 db.Close
 Set changes = CreateObject("Scripting.Dictionary"): changes("Trait33Comment1") = "LOCAL CONFLICT VALUE"
 Call CommitSavedFields(id, CurrentRevision(id), changes, "Synthetic three-way comment conflict")
 Set rows = NAVFITPreviewReturned(returned, file & ".swift-baseline.accdb")
 Check NAVFITReturnedDecision("Trait33Comment1") = "Conflict", "Chief authored conflict was not surfaced"
 NAVFITResolveReturned "Trait33Comment1", "Keep current": expected("Trait33Comment1") = "LOCAL CONFLICT VALUE"
 Check NAVFITApplyReturned() = id, "Chief comment return failed"
 Set fields = SavedFields(id)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check SameFieldValue(fields(key), expected(key)), "Returned comment lost or crossed: " & key
  Next
 Next
 Set changes = CreateObject("Scripting.Dictionary"): changes("Trait33Comment2") = Null
 Call CommitSavedFields(id, CurrentRevision(id), changes, "Deliberately clear authored Chief comment")
 file = root & "\chief-cleared.accdb": Call NAVFITExportReport(id, file, template)
 Set db = DBEngine.OpenDatabase(file, False, True): Set native = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
 Check IsNull(native!LEADDN2), "Explicit comment clear was replaced by native provenance"
 native.Close: db.Close
 Check RestoreDraft(id, authored) = id, "Restoring authored Chief revision failed"
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check SavedFields(id)(key) = "AUTHORED " & block & "." & entry & " O'Brien & evidence", "Restore omitted authored comment: " & key
  Next
 Next
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Check UXFixedSelect(0), "Chief Front Page cannot open"
 Set surface = Forms!frmEditor!uxFieldSurface.Form: revision = CurrentRevision(id)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry: text = "PENDING " & block & "." & entry
   surface.Controls(UXFieldControl(key)) = text
   Check UXSurfaceEdited(key), "Chief comment edit failed: " & key
   expected(key) = text
  Next
 Next
 Check CheckpointEditor(), "Chief comment recovery checkpoint failed"
 recovery = DMax("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check DLookup(key, "RecoveryDraft", "RecoveryDraftID=" & recovery) = expected(key), "Recovery omitted " & key
  Next
 Next
 Check UXFixedSelect(5) And UXFixedSelect(0), "Chief comments cannot cross Back/Front pages"
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check Forms!frmEditor.Controls(UXFieldControl(key)) = expected(key), "Page switch lost " & key
  Next
 Next
 Check CurrentRevision(id) = revision, "Pending comment navigation saved implicitly"
 Check UXSaveEditor(), "Chief authored comment save failed"
 Check UXLeaveEditor(), "Chief editor cannot close after save"
 Set fields = SavedFields(id)
 For block = 33 To 39
  For entry = 1 To 3
   key = "Trait" & block & "Comment" & entry
   Check SameFieldValue(fields(key), expected(key)), "Saved editor comment differs: " & key
  Next
 Next
 Set changes = CreateObject("Scripting.Dictionary"): changes("Trait39Comment3") = String$(800, "W") & " END_CHIEF_39_3"
 Call CommitSavedFields(id, CurrentRevision(id), changes, "Synthetic full overflow preservation")
 revision = CurrentRevision(id): file = root & "\chief-too-long.accdb"
 Check InStr(ChiefExportFailure(id, file, template), "exceeds native capacity") > 0, "Native 255-character capacity was silently coerced"
 Check Len(Dir$(file)) = 0 And CurrentRevision(id) = revision, "Failed native export published or changed report"
 PDFOutputCapture 0, id, False: Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True
 folder = PDFOutputGenerate(root, selected, "Individual PDFs", False)
 Check Len(folder) > 0 And SavedFields(id)("Trait39Comment3") = changes("Trait39Comment3"), "PDF overflow changed full comment"
 TestChiefComments = "PASS: " & n - start & " Chief authored-field, native-return, recovery, editor and overflow assertions; " & folder: Exit Function
Failed:
 TestChiefComments = "FAIL Chief comments after " & n - start & ": " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
Private Function ChiefExportFailure(ByVal id As Long, ByVal file As String, ByVal template As String) As String
 On Error GoTo Failed
 Call NAVFITExportReport(id, file, template): ChiefExportFailure = "ACCEPTED": Exit Function
Failed:
 ChiefExportFailure = Err.Description
End Function
Public Function TestChiefPrintCases(ByVal root As String) As String
 Dim groupID As Long, id As Long, which As Long, block As Long, entry As Long, mark As Long, marks As Variant, name As String, key As String, text As String, folder As String
 Dim changes As Object, selected As Object, current As Object, rows As Collection
 On Error GoTo Failed
 If DCount("*", "UserProfile") > 0 Then BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
 ConfigureWorkspace "SYNTHETIC CHIEF PRINT CASES", "Synthetic tester"
 groupID = CreateSummaryGroup("Synthetic Chief print matrix", "CHIEFEVAL", "E7")
 For which = 0 To 8
  Select Case which
   Case 0: name = "ZULU": mark = 3
   Case 1: name = "CHARLIE": mark = 4
   Case 2: name = "ALPHA": mark = 2
   Case 3: name = "BRAVO": mark = 5
   Case Else: mark = which - 3: name = "GRADE" & mark
  End Select
  marks = Array(mark & ".0", mark & ".0", mark & ".0", mark & ".0", mark & ".0", mark & ".0", mark & ".0")
  id = CommitDraft(0, "CHIEFEVAL", "Synthetic " & name, name & ", SYNTHETIC", #1/1/2026#, #9/30/2026#, "", "Synthetic print matrix", "Synthetic reporting senior narrative.", , marks, Array("Observed", "Promotable"))
  Set changes = CreateObject("Scripting.Dictionary")
  changes("OfficialName") = name & ", SYNTHETIC": changes("RankRate") = "YNC"
  changes("ReportingSeniorName") = "SYNTHETIC, A B": changes("ReportingSeniorGrade") = "CAPT"
  changes("MemberSSN") = "000000" & Format$(which, "000"): changes("UnitUIC") = "00001"
  For block = 33 To 39
   For entry = 1 To 3
    key = "Trait" & block & "Comment" & entry
    Select Case which
     Case 0: text = "OVERFLOW " & block & "." & entry & " " & String$(500, "W") & " END_OVERFLOW_" & block & "_" & entry
     Case 1: text = "FULL " & block & "." & entry & " first authored line of evidence." & vbCrLf & "FULL " & block & "." & entry & " second authored line of evidence."
     Case 2: text = ""
     Case 3: text = "SHORT " & block & "." & entry & " distinct evidence."
     Case Else: text = ""
    End Select
    If text <> "" Then changes(key) = text
   Next
  Next
  Call CommitSavedFields(id, CurrentRevision(id), changes, "Synthetic print case values")
  AddGroupMember groupID, id, "E7"
 Next
 PDFOutputCapture groupID, 0, False: Set current = PDFOutputJob(): Set rows = current("Reports")
 Check rows.Count = 9 And rows(1)("Last") = "ALPHA" And rows(2)("Last") = "BRAVO" And rows(3)("Last") = "CHARLIE" And rows(9)("Last") = "ZULU", "Chief matrix packet order differs"
 Set selected = CreateObject("Scripting.Dictionary")
 For which = 1 To rows.Count: selected(CStr(rows(which)("ID"))) = True: Next
 folder = PDFOutputGenerate(root, selected, "Both", True)
 Check current("Files").Count = 11, "Chief matrix omitted packet, summary or an individual report"
 TestChiefPrintCases = "PASS: blank/short/two-line/overflow comments, all 35 grade combinations, summary and alphabetical Both output: " & folder: Exit Function
Failed:
 TestChiefPrintCases = "FAIL: Chief print matrix: " & Err.Number & " " & Err.Description
End Function
Public Function TestPDFOutput(ByVal root As String) As String
 Dim g As Long, id As Long, other As Long, fields As Object, current As Object, rows As Collection, selected As Object, folder As String, again As String, before As Long, key As Variant
 On Error GoTo Failed
 If DCount("*", "UserProfile") > 0 Then BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
 ConfigureWorkspace "SYNTHETIC PDF OUTPUT", "Tester"
 Check PDF_OUTPUT_SLOTS >= UBound(PDFSummaryRows(2)) * 6 + 22, "Renderer cannot cover every populated summary continuation cell"
 g = CreateSummaryGroup("Synthetic PDF Group", "EVAL", "E5")
 id = CommitDraft(0, "EVAL", "Zulu draft", "ZULU, TEST", Null, Null, "")
 other = CommitDraft(0, "EVAL", "Alpha draft", "ALPHA, TEST", Null, Null, "")
 AddGroupMember g, id, "E5": AddGroupMember g, other, "E5"
 before = DCount("*", "ReportRevision")
 PDFOutputCapture g, 0, False: Set current = PDFOutputJob(): Set rows = current("Reports")
 Check rows.Count = 2, "Capture omitted incomplete reports"
 Check rows(1)("Last") = "ALPHA" And rows(2)("Last") = "ZULU", "Last-name order incorrect"
 Check DCount("*", "ReportRevision") = before, "Capture changed revisions"
 Check PDFOutputSafePart("123-45-6789 1234567890 / CON") = "ID_ID___CON", "Filename redaction or sanitization failed"
 Check PDFOutputSafePart("CON") = "_CON", "Reserved filename unprotected"
 'Block 48 holds the address only; the phone is for the summary letter. An unticked box with
 'no place on the form, or a field this report type no longer has, is not added as extra text.
 Set fields = rows(2)("Fields")
 fields("ReportingSeniorAddress1") = "1 SYNTHETIC WAY": fields("ReportingSeniorPhone") = "5550001234"
 fields("ReportOperationalCommander") = "No": fields("ReadinessDate") = "2026-11-01"
 Check InStr(PDFOutputPlannedText(rows(2), False), "1 SYNTHETIC WAY") > 0, "Block 48 address missing"
 Check InStr(PDFOutputPlannedText(rows(2), False), "5550001234") = 0, "Phone printed on the report"
 again = PDFOutputPlannedText(rows(2), True)
 Check InStr(again, UXFieldLabel("ReportOperationalCommander")) = 0 And InStr(again, "2026-11-01") = 0 And InStr(again, "11/26") = 0, "Unticked or retired field added as extra text: " & again
 fields("ReportOperationalCommander") = "Yes"
 Check InStr(PDFOutputPlannedText(rows(2), True), UXFieldLabel("ReportOperationalCommander")) > 0, "Ticked box without a place on the EVAL was dropped"
 fields("ReportOperationalCommander") = Null: fields("ReadinessDate") = Null: again = ""
 'Full dates print as YYMMMDD on every report type and in summary letter block 12.
 fields("PeriodStart") = DateSerial(2026, 5, 16): fields("PeriodEnd") = DateSerial(2026, 11, 7): fields("DateReported") = "2025-01-31"
 rows(1)("Fields")("PeriodEnd") = DateSerial(2026, 11, 7)
 again = PDFOutputPlannedText(rows(2), False)
 Check InStr(again, "26MAY16") > 0 And InStr(again, "26NOV07") > 0 And InStr(again, "25JAN31") > 0 And InStr(again, "2026-11-07") = 0, "Report dates not printed as YYMMMDD"
 Check InStr(PDFOutputPlannedText(Nothing, False), "26NOV07") > 0, "Summary letter end date not printed as YYMMMDD"
 'Block 44 prints in its EVAL box; text past the box continues on an added page.
 fields("Qualifications") = "SYNTHETIC BLOCK 44 COURSE"
 Check InStr(PDFOutputPlannedText(rows(2), False), "SYNTHETIC BLOCK 44 COURSE") > 0 And InStr(PDFOutputPlannedText(rows(2), True), "SYNTHETIC BLOCK 44 COURSE") = 0, "Block 44 not printed in its box"
 fields("Qualifications") = "BLOCK44 LINE 1" & vbCrLf & "BLOCK44 LINE 2" & vbCrLf & "BLOCK44 LINE 3" & vbCrLf & "BLOCK44 LINE 4" & vbCrLf & "BLOCK44 LINE 5" & vbCrLf & "BLOCK44 LAST"
 again = PDFOutputPlannedText(rows(2), True)
 Check InStr(again, "BLOCK44 LAST") > 0 And InStr(again, "BLOCK44 LINE 1") = 0, "Block 44 overflow not continued on an added page: " & again
 fields("Qualifications") = Null
 again = ""
 'Synthetic renderer stress values do not alter saved reports.
 Set fields = rows(1)("Fields")
 fields("OfficialName") = "ALPHA, TEST A": fields("OfficialComments") = "OVERFLOW START " & String$(16000, "W") & " OVERFLOW END"
 fields("PeriodEnd") = "incomplete": fields("MemberSSN") = "000-00-0123"
 fields("ReportingSeniorGrade") = "LONG-GRADE-TAIL"
 For Each key In TraitKeys(): fields(CStr(key)) = "3.0": Next
 Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True: selected(CStr(other)) = True
 folder = PDFOutputGenerate(root, selected, "Both", True)
 Check CreateObject("Scripting.FileSystemObject").FolderExists(folder), "Completed folder absent"
 Check current("Files").Count = 4, "Combined, summary and individual inventory incorrect"
 Check Not CreateObject("Scripting.FileSystemObject").FileExists(folder & "\pages.accdb"), "Transient page cache retained after success"
 Check DCount("*", "ReportRevision") = before, "Export changed report revisions"
 again = PDFOutputGenerate(root, selected, "Individual PDFs", False)
 Check folder <> again, "Repeat export overwrote a run"
 Check current("Files").Count = 2, "Report-only output inventory incorrect"
 TestPDFOutput = "PASS: " & n & " PDF capture, ordering, nonblocking output and isolation assertions; " & folder: Exit Function
Failed:
 TestPDFOutput = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
Private Function PDFFilePageCount(ByVal path As String) As Long
 Dim f As Integer, data As String, at As Long
 f = FreeFile: Open path For Binary Access Read As #f: data = Space$(LOF(f)): Get #f, , data: Close #f
 at = InStr(1, data, "/Type/Page", vbBinaryCompare)
 Do While at > 0
  If Mid$(data, at + 10, 1) <> "s" Then PDFFilePageCount = PDFFilePageCount + 1
  at = InStr(at + 10, data, "/Type/Page", vbBinaryCompare)
 Loop
End Function
Private Function PacketPageCount(ByVal folder As String) As Long
 Dim file As Variant
 For Each file In PDFOutputJob()("Files")
  If Right$(file, 12) = "__Packet.pdf" Then PacketPageCount = PDFFilePageCount(folder & "\" & file)
 Next
End Function
'Pages print edge to edge. With a physical printer as Access's printer, Access spread each
'page over several and warned that the section was wider than the page (0.55.0 on a
'workstation whose default printer was a network printer). Output must not depend on it.
Public Function TestPDFLayoutPrinter(ByVal root As String) As String
 Dim g As Long, id As Long, other As Long, selected As Object, p As Access.Printer, layout As String, margined As String, expected As Long, folder As String
 On Error GoTo Failed
 layout = PDFLayoutPrinterName()
 For Each p In Application.Printers
  If p.DeviceName = "Microsoft Print to PDF" Then Check layout = p.DeviceName, "Microsoft Print to PDF is installed but not used for layout: " & layout
 Next
 For Each p In Application.Printers
  If p.DeviceName <> layout And InStr(1, p.DeviceName, "PDF", vbTextCompare) = 0 And InStr(1, p.DeviceName, "XPS", vbTextCompare) = 0 And InStr(1, p.DeviceName, "Fax", vbTextCompare) = 0 Then margined = p.DeviceName: Exit For
 Next
 If DCount("*", "UserProfile") > 0 Then BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
 ConfigureWorkspace "SYNTHETIC PDF PRINTER", "Tester"
 g = CreateSummaryGroup("Synthetic Printer Group", "EVAL", "E5")
 id = CommitDraft(0, "EVAL", "Printer draft one", "ONE, TEST", Null, Null, "")
 other = CommitDraft(0, "EVAL", "Printer draft two", "TWO, TEST", Null, Null, "")
 AddGroupMember g, id, "E5": AddGroupMember g, other, "E5"
 PDFOutputCapture g, 0, False
 Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True: selected(CStr(other)) = True
 Set Application.Printer = Application.Printers(layout)
 folder = PDFOutputGenerate(root, selected, "Combined packet", True): expected = PacketPageCount(folder)
 Check expected > 0, "Packet page count unreadable"
 Check Application.Printer.DeviceName = layout, "Access printer not left as it was"
 If margined <> "" Then
  Set Application.Printer = Application.Printers(margined)
  folder = PDFOutputGenerate(root, selected, "Combined packet", True)
  Check PacketPageCount(folder) = expected, "Pages depend on Access's printer (" & margined & "): " & PacketPageCount(folder) & " pages, expected " & expected
  Check Application.Printer.DeviceName = margined, "Access printer not restored after PDF output"
 End If
 Set Application.Printer = Nothing
 TestPDFLayoutPrinter = "PASS: " & n & " layout printer assertions; " & expected & "-page packet laid out on " & layout & IIf(margined = "", "; no other printer installed to compare", "; same pages with " & margined & " as Access's printer"): Exit Function
Failed:
 Set Application.Printer = Nothing
 TestPDFLayoutPrinter = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
Public Function TestPDFBoundaries(ByVal root As String) As String
 Dim g As Long, id As Long, current As Object, seed As Object, row As Object, fields As Object, rows As New Collection, selected As Object, key As Variant, i As Long, folder As String, other As Object, before As Long
 On Error GoTo Failed
 g = CreateSummaryGroup("Boundary Summary", "EVAL", "E5")
 PDFOutputCapture g, 0, False: Set current = PDFOutputJob()
 Set selected = CreateObject("Scripting.Dictionary")
 folder = PDFOutputGenerate(root, selected, "Combined packet", True)
 Check current("Files").Count = 1, "Empty group summary export failed"
 id = CommitDraft(0, "EVAL", "Archived PDF", "ARCHIVED, PRINT", Null, Null, "")
 Check ArchiveDraft(id, True), "Archive fixture setup failed"
 before = CurrentRevision(id)
 PDFOutputCapture 0, id, False: Set current = PDFOutputJob(): Set seed = current("Reports")(1)
 selected(CStr(id)) = True: folder = PDFOutputGenerate(root, selected, "Individual PDFs", False)
 Check CurrentRevision(id) = before And DLookup("IsArchived", "Report", "ReportID=" & id), "Archived output changed state"
 Check PDFOutputSafePart("name1234567890suffix") = "nameIDsuffix", "Embedded identifier exposed in filename"
 'Plain-VBA matching (no VBScript); expected values are the former RegExp results.
 Check PDFOutputSafePart("123 45 6789") = "ID" And PDFOutputSafePart("12345678901") = "ID1", "Spaced or eleven-digit identifier exposed"
 Check PDFOutputSafePart("123-456-7890") = "123-456-7890", "Phone-shaped number changed"
 Check PDFOutputSafePart("lpt1.txt") = "_lpt1.txt" And PDFOutputSafePart("COM0") = "COM0" And PDFOutputSafePart("Console") = "Console", "Device-name protection differs"
 Set other = CreateObject("Scripting.Dictionary")
 For Each key In Array("Last", "First", "Middle", "Guid"): other(CStr(key)) = seed(key): Next
 seed("Last") = "SAME": seed("First") = "ALPHA": seed("Middle") = "A": seed("Guid") = "A"
 other("Last") = "SAME": other("First") = "BETA": other("Middle") = "A": other("Guid") = "A"
 Check PDFOutputCompare(seed, other) < 0, "First-name tie break failed"
 other("First") = "ALPHA": other("Middle") = "B": Check PDFOutputCompare(seed, other) < 0, "Middle-name tie break failed"
 other("Middle") = "A": other("Guid") = "B": Check PDFOutputCompare(seed, other) < 0, "Stable identity tie break failed"
 PDFOutputCapture g, 0, False: Set current = PDFOutputJob(): Set selected = CreateObject("Scripting.Dictionary")
 'Renderer-only roster stress snapshot: fill both columns of a continuation sheet.
 For i = 1 To 144
  Set row = CreateObject("Scripting.Dictionary"): Set fields = CreateObject("Scripting.Dictionary")
  For Each key In seed.Keys
   If key <> "Fields" Then row.Add CStr(key), seed(key)
  Next
  fields.Add "FormKind", "EVAL": fields.Add "OfficialName", "ROSTER" & Format$(i, "000") & ", TEST"
  fields.Add "MemberSSN", "000-00-" & Format$(i, "0000")
  For Each key In TraitKeys(): fields.Add CStr(key), "3.0": Next
  row.Add "Fields", fields: row("ID") = -i: row("Guid") = NewIdentity()
  row("Last") = "ROSTER" & Format$(i, "000"): row("Caption") = fields("OfficialName")
  rows.Add row: selected(CStr(-i)) = True
 Next
 Set current("Reports") = rows: Set current("StatisticsRows") = rows
 folder = PDFOutputGenerate(root, selected, "Combined packet", True)
 Check current("Reports").Count = 144 And current("Files").Count = 1, "Summary continuation roster omitted reports"
 TestPDFBoundaries = "PASS: " & n & " cumulative assertions including empty/archived output, tie breaks and fully populated 144-member summary continuation; " & folder: Exit Function
Failed:
 TestPDFBoundaries = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
Private Function BadDestination(ByVal selected As Object, ByVal root As String) As Boolean
 On Error GoTo Expected
 PDFOutputGenerate root & "\missing_" & NewIdentity(), selected, "Both", True
 Exit Function
Expected:
 BadDestination = (Err.Number = 76)
End Function
Public Function TestPDFDraftsAndKinds(ByVal root As String) As String
 Dim id As Long, g As Long, before As Long, current As Object, row As Object, selected As Object, fields As Object, key As Variant, kind As Variant, folder As String, grid As Access.Form, value As String, state As String, stage As String, countRows As Collection, countIndex As Long, isicID As Long
On Error GoTo Failed
 If DCount("*", "UserProfile") > 0 Then BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
 id = CommitDraft(0, "EVAL", "Pending PDF", "PENDING, EDITOR", Null, Null, "")
 before = CurrentRevision(id)
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Forms!frmEditor!txtOfficialComments = "EDITOR PENDING PDF SENTINEL"
 PDFOutputCapture 0, id, True: Set current = PDFOutputJob(): Set row = current("Reports")(1)
 Check row("Fields")("OfficialComments") = "EDITOR PENDING PDF SENTINEL", "Editor pending values were not captured"
 Check CurrentRevision(id) = before And EditorHasPending(), "Capture saved or discarded editor draft"
 Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True
 folder = PDFOutputGenerate(root, selected, "Individual PDFs", False)
 Check CurrentRevision(id) = before And EditorHasPending(), "PDF generation saved or discarded editor draft"
 Check PDFOutputOpenReport(), "Report PDF entry failed with pending text"
 Check Forms!frmPDFOutput!lstReports.ListCount = 1 And Forms!frmPDFOutput!cmdGenerate.Enabled, "Report preview inventory or nonblocking action unavailable"
 Check Forms!frmPDFOutput!cmdPagePreview.OnClick = "=PDFOutputPagePreview()", "Native page-preview binding missing"
 Forms!frmPDFOutput!txtFolder = root: Forms!frmPDFOutput!cmdGenerate.SetFocus
 Check PDFOutputCreate(), "Focused Generate button callback failed: " & Forms!frmPDFOutput!lblStatus.Caption
 Check InStr(Forms!frmPDFOutput!txtFindings, PDFOutputJob()("Folder")) > 0 And InStr(Forms!frmPDFOutput!txtFindings, PDFOutputJob()("Files")(1)) > 0, "Output folder path and files not shown for copying"
 Check CurrentRevision(id) = before And EditorHasPending(), "UI export changed pending editor state"
 Check PDFOutputClose(), "Report PDF preview close failed"
 'Preview in the editor shows the PDF pages with pending edits, saving nothing.
 PDFOutputRememberFolder root
 Check UXEditorPrintPreview(), "Editor print preview failed: " & Forms!frmEditor!uxError.Caption
 Check CurrentProject.AllReports("rptPDFOutput").IsLoaded, "Editor print preview not shown"
 Check InStr(PDFOutputPlannedText(PDFOutputJob()("Reports")(1), False), "EDITOR PENDING PDF SENTINEL") > 0, "Print preview missed pending text"
 Check CurrentRevision(id) = before And EditorHasPending(), "Print preview saved or discarded editor draft"
 PDFOutputEndPreview
 Check Not CurrentProject.AllReports("rptPDFOutput").IsLoaded, "Print preview did not close"
 DeleteSetting "SwiftNAVFIT", "PDF-" & MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1")
 PDFOutputCapture 0, id, False: Set current = PDFOutputJob(): Set row = current("Reports")(1)
 Check Nz(row("Fields")("OfficialComments"), "") <> "EDITOR PENDING PDF SENTINEL", "Saved mode read pending editor buffer"
 Check EditorHasPending(), "Saved capture cleared pending editor"
 Check UXLeaveEditor("Discard"), "Synthetic editor cleanup failed"
 g = CreateSummaryGroup("Pending grid PDF", "EVAL", "E5"): AddGroupMember g, id, "E5"
 Check UXShowGroupMembers(g, True), "Grid setup failed"
 Set grid = Forms!frmUXComparison!uxTraitGrid.Form
 value = String$(21, "X"): grid!AssignmentRecommendationA = value
 Check UXGridQueue("AssignmentRecommendationA"), "Pending queue setup failed"
 Forms!frmUXComparison.TimerInterval = 0
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Forms!frmEditor!txtOfficialComments = "COEXISTING EDITOR DRAFT"
 PDFOutputCapture g, 0, True: Set current = PDFOutputJob(): Set row = current("Reports")(1)
 Check row("Fields")("AssignmentRecommendationA") = value, "Queued invalid cell absent from snapshot"
 Check row("Fields")("OfficialComments") = "COEXISTING EDITOR DRAFT" And EditorHasPending(), "Group capture replaced another editor field"
 Check UXGridPendingCount() = 1 And CurrentRevision(id) = before, "Capture drained pending cells"
 isicID = ISICSave(0, 0, "Synthetic picker ISIC " & NewIdentity(), "N00001", "Synthetic ISIC address")
 Check PDFOutputOpenGroup("frmUXComparison"), "Group PDF entry failed with pending cells"
 With Forms!frmPDFOutput
  Check !cboISIC.RowSourceType = "Table/Query" And !cboISIC.ListCount = DCount("*", "ISIC", "IsArchived=False"), "ISIC choices not listed from the ISIC table"
  Check !cboOnTime.ListCount = 2 And !lblReports.Top >= !cboISIC.Top + !cboISIC.Height And !lblReports.Top >= !cboOnTime.Top + !cboOnTime.Height, "Report-order caption overlaps the summary-letter choices"
  'Choose the ISIC as a user does, by its list row: the picker must return the ISIC ID.
  Check !cboISIC.BoundColumn = 1, "ISIC picker is not bound to the ISIC ID"
  For countIndex = 0 To !cboISIC.ListCount - 1
   If CLng(!cboISIC.Column(0, countIndex)) = isicID Then !cboISIC = !cboISIC.ItemData(countIndex)
  Next
  Check Nz(!cboISIC, "") = CStr(isicID), "Choosing an ISIC from the list did not return its ID"
  !cboOnTime = !cboOnTime.ItemData(0)
 End With
 Forms!frmPDFOutput!txtFolder = root
 Check PDFOutputPagePreview(), "Group native page preview failed"
 Check CurrentProject.AllReports("rptPDFOutput").IsLoaded And UXGridPendingCount() = 1, "Page preview lost pending grid cells"
 PDFOutputEndPreview
 Check PDFOutputClose(), "Group PDF preview close failed"
 Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True
 folder = PDFOutputGenerate(root, selected, "Combined packet", True)
 Check UXGridPendingCount() = 1 And CurrentRevision(id) = before, "Export changed pending cells or saved revision"
 Check BadDestination(selected, root) And Not PDFOutputBusy(), "I/O failure did not preserve error/release busy state"
 Check UXLeaveEditor("Discard"), "Synthetic coexisting editor cleanup failed"
 Check UXGridDiscardPending(), "Synthetic queue cleanup failed"
 Check UXCloseComparison(), "Synthetic grid cleanup failed"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  stage = "Create kind fixture"
  id = CommitDraft(0, CStr(kind), "Synthetic print " & kind, "PRINT, " & kind, Null, Null, "", _
      reason:="PDF saved-field readback proof", reportingSenior:=Array("", "", "", "", "54321"), _
      serviceDetails:=Array("", "", "PASS", "", "2026-06-01", "BASIC"), _
      reportPeople:=Array("000009876", "", "", "", "", "", "", "", "", "", "", "", ""))
  stage = "Capture saved-field proof"
  PDFOutputCapture 0, id, False: Set current = PDFOutputJob(): Set row = current("Reports")(1): Set fields = row("Fields")
  stage = "Check saved-field proof"
  Check fields("ReportingSeniorSSN") = "000009876" And fields("ReportingSeniorUIC") = "54321" And fields("ReadinessDate") = "2026-06-01" And fields("BilletSubcategory") = "BASIC", "Saved-field capture changed values: " & kind
  For Each key In DraftFieldKeys()
   If PDFPrintAdditionalField(CStr(key)) Then fields(CStr(key)) = "PROOF"
  Next
  fields("FormKind") = kind: fields("OfficialName") = "PRINT, " & kind: fields("MemberSSN") = "000-00-0000"
  fields("PeriodStart") = "2026-01-01": fields("PeriodEnd") = "2026-09-30"
  fields("ReportingSeniorSSN") = "000009876": fields("ReportingSeniorUIC") = "54321"
  fields("ReportingSeniorGrade") = "LCDR": fields("ReportingSeniorDesignator") = "1110"
  fields("ReadinessDate") = "2026-06-01": fields("ReadinessSummary") = "PASS": fields("ReadinessDetail") = ""
  fields("BilletSubcategory") = "BASIC": fields("ObservationStatus") = "Observed"
  fields("Trait33") = "1.0": fields("Trait34") = "2.0": fields("Trait35") = "3.0": fields("Trait36") = "4.0"
  fields("Trait37") = "5.0": fields("Trait38") = "NOB": fields("Trait39") = "3.0"
  For Each key In Array("ComponentActive", "OccasionPeriodic", "ReportRegular", "RetentionYes", "StatementNo"): fields(CStr(key)) = "Yes": Next
  fields("PromotionRecommendation") = "Significant Problems"
  If kind = "CHIEFEVAL" Then
   'Block 20: a CHIEFEVAL's result is edited as ReadinessDetail and prints beside MM/YY.
   fields("ReadinessDetail") = "PPPB"
   Check InStr(PDFOutputPlannedText(row, False), "PPPB") > 0, "CHIEFEVAL block 20 result not printed"
   'Renderer-only count stress, preserving the saved report and membership tables.
   Set countRows = New Collection
   For countIndex = 1 To 144: countRows.Add row: Next
   current("StatisticsGroupID") = 1: Set current("StatisticsRows") = countRows
  End If
  Set selected = CreateObject("Scripting.Dictionary"): selected(CStr(id)) = True
  before = CurrentRevision(id): state = Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "")
  folder = PDFOutputGenerate(root, selected, "Individual PDFs", False)
  Check current("Files").Count = 1 And CurrentRevision(id) = before, "Kind output changed revisions: " & kind
  Check Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") = state, "Export changed workflow: " & kind
  PDFOutputPreviewPages root, False
  Check CurrentProject.AllReports("rptPDFOutput").IsLoaded, "Native page preview unavailable"
  PDFOutputEndPreview
 Next
 TestPDFDraftsAndKinds = "PASS: " & n & " cumulative pending-editor/grid, failure, report-kind and native page-preview assertions": Exit Function
Failed:
 TestPDFDraftsAndKinds = "FAIL after " & n & " (" & Nz(kind, "pending") & "/" & stage & "): " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
