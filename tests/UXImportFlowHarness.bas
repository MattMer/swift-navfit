Attribute VB_Name = "UXImportFlowHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5100, , message
    checks = checks + 1
End Sub
Public Function TestUXImportFlow(ByVal source As String, ByVal pdfPath As String, Optional ByVal includePDF As Boolean = True) As String
    Dim before As Long, revision As Long, text As String, rejected As Boolean, id As Long, replacement As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC IMPORT FLOW", "Synthetic tester"
    Check OpenReviewFile(source), "Source choose failed"
    Check Not Forms!frmImportReview!cmdImport.Enabled, "Apply enabled before preview"
    Check Forms!frmImportReview.OnUnload = "=CloseReviewFile()", "Native source close does not release reader"
    Check Forms!frmImportReview!cmdPreview.TabIndex < Forms!frmImportReview!cmdImport.TabIndex, "Narrative tab order skips preview"
    before = DCount("*", "Report")
    Forms!frmImportReview!lstImportReports = 1: Call ReviewReportFields
    Check Not UXPreviewNarrativeImport(), "Quarantined source preview accepted"
    Forms!frmImportReview!lstImportReports = 2: Call ReviewReportFields
    Forms!frmImportReview!txtImportMember = "Synthetic import member"
    On Error Resume Next
    id = ImportReviewedNarrative(3, "Synthetic stale source", 0, 0, String$(64, "0"))
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    Check rejected And DCount("*", "Report") = before, "Changed-source guard wrote a report"
    Check UXPreviewNarrativeImport(), "Narrative preview failed"
    Check DCount("*", "Report") = before, "Preview wrote reports"
    Check InStr(Forms!frmImportReview!uxPreview, "Before: [new draft") > 0 And InStr(Forms!frmImportReview!uxPreview, "After:") > 0, "Before/after values absent"
    Check InStr(Forms!frmImportReview!uxPreview, "unmapped") > 0, "Mapping scope absent"
    Check UXNarrativeInspect() And Forms!frmImportReview!lstImportFields.Visible, "Raw inspection lost"
    Forms!frmImportReview!txtImportMember = "Changed member"
    Check Not UXApplyNarrativeImport(), "Changed unpreviewed member accepted"
    Check DCount("*", "Report") = before And Not Forms!frmImportReview!cmdImport.Enabled, "Stale Apply wrote report"
    Check Forms!frmImportReview!txtImportMember = "Changed member", "Failure lost pending input"
    Check UXPreviewNarrativeImport(), "Repreview failed"
    Check UXApplyNarrativeImport(), "Narrative apply failed"
    Check DCount("*", "Report") = before + 1 And Not Forms!frmImportReview!cmdImport.Enabled, "Apply results state wrong"
    Check InStr(Forms!frmImportReview!uxState.Caption, "4. Results") > 0, "Results missing"
    Check Not UXApplyNarrativeImport() And DCount("*", "Report") = before + 1, "Repeated apply wrote draft"
    Check UXOpenImportedReport("NAVFIT"), "Imported report cannot open"
    Check Forms!frmEditor!txtMember = "Changed member", "Wrong imported report opened"
    Call UXLeaveEditor
    Call CloseReviewFile
    If Not includePDF Then TestUXImportFlow = "PASS: " & checks & " NAVFIT guided import assertions; PDF integration not requested": Exit Function
    DoCmd.OpenForm "frmPDFImport"
    before = DCount("*", "Report")
    Check Not Forms!frmPDFImport!cmdImport.Enabled, "PDF Apply enabled before choose"
    Check Forms!frmPDFImport.OnUnload = "=PDFClearPreview(False)", "Native PDF close retains preview"
    Check Forms!frmPDFImport!cmdPreview.TabIndex < Forms!frmPDFImport!cmdImport.TabIndex, "PDF tab order skips preview"
    Check InStr(Forms!frmPDFImport!cboIdentifier.RowSource, "DoDID") = 0, "Legacy ID exposed by default"
    Check UXChoosePDFFile(pdfPath), "PDF choose failed"
    Forms!frmPDFImport!cboIdentifier = "SSN"
    Check UXPreviewPDFImport(), "PDF mapping preview failed"
    Check InStr(Forms!frmPDFImport!txtPreview, "4. SSN") > 0 And InStr(Forms!frmPDFImport!txtPreview, "After: 000000000") > 0, "Official caption/value absent"
    Check DCount("*", "Report") = before, "PDF preview wrote draft"
    Set replacement = PreviewStructuredPDF(Replace(pdfPath, "synthetic-eval.pdf", "synthetic-fitrep.pdf"))
    Check Not UXApplyPDFImport() And DCount("*", "Report") = before, "Changed PDF bypassed preview"
    Check UXChoosePDFFile(pdfPath) And UXPreviewPDFImport(), "PDF restore/repreview failed"
    Forms!frmPDFImport!cboIdentifier = "DoDID"
    Check Not UXPreviewPDFImport(), "Implicit legacy mapping accepted"
    Forms!frmPDFImport!chkLegacy = True: Call UXPDFLegacyChanged
    Check UXPreviewPDFImport(), "Explicit legacy compatibility preview failed"
    Forms!frmPDFImport!chkLegacy = False: Call UXPDFLegacyChanged
    Check Forms!frmPDFImport!cboIdentifier = "SSN" And Not Forms!frmPDFImport!cmdImport.Enabled, "Leaving legacy mode retained stale Apply"
    Check UXPreviewPDFImport(), "SSN preview restore failed"
    Forms!frmPDFImport!cboIdentifier = "Retain only"
    Check Not UXApplyPDFImport() And DCount("*", "Report") = before, "Changed identifier bypassed preview"
    Forms!frmPDFImport!cboIdentifier = "SSN": Check UXPreviewPDFImport(), "PDF repreview failed"
    Check UXApplyPDFImport(), "PDF Apply failed: " & Forms!frmPDFImport!uxState.Caption
    Check InStr(Forms!frmPDFImport!uxState.Caption, "4. Results") > 0 And Forms!frmPDFImport!cmdOpenResult.Enabled, "PDF results missing"
    Check Not UXApplyPDFImport() And DCount("*", "Report") = before + 1, "PDF repeated Apply created draft"
    Check UXChoosePDFFile(pdfPath), "Reopen test could not choose source"
    DoCmd.Close acForm, "frmPDFImport", acSaveNo
    DoCmd.OpenForm "frmPDFImport"
    Check Not UXPreviewPDFImport() And Not Forms!frmPDFImport!cmdImport.Enabled, "Reopening retained hidden PDF preview"
    TestUXImportFlow = "PASS: " & checks & " guided import preview/state/results assertions": Exit Function
Failed:
    TestUXImportFlow = "FAIL after " & checks & ": " & Err.Description
End Function
