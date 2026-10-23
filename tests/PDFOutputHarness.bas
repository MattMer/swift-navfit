Attribute VB_Name = "PDFOutputHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 9190, , message
 n = n + 1
End Sub
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
 Dim id As Long, g As Long, before As Long, current As Object, row As Object, selected As Object, fields As Object, key As Variant, kind As Variant, folder As String, grid As Access.Form, value As String, state As String, stage As String, countRows As Collection, countIndex As Long
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
 Check PDFOutputCreate(), "Focused Generate button callback failed"
 Check CurrentRevision(id) = before And EditorHasPending(), "UI export changed pending editor state"
 Check PDFOutputClose(), "Report PDF preview close failed"
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
 Check PDFOutputOpenGroup("frmUXComparison"), "Group PDF entry failed with pending cells"
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
