Attribute VB_Name = "PDFCounselingStatusHarness"
Option Explicit
Public Function TestPDFCounselingStatuses(ByVal directory As String) As String
    Dim kind As Variant, status As Variant, preview As Object, fields As Object, saved As Object, id As Long, n As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC COUNSELING IMPORT", "Status tester"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        For Each status In Array("NOT REQ", "NOT PERF")
            Set preview = PreviewStructuredPDF(directory & "\" & LCase$(kind) & "-" & Replace(status, " ", "-") & ".pdf")
            Set fields = PDFMappedValues("20", "SSN")
            If fields("CounselingDate") <> status Then Err.Raise vbObjectError + 4751, , "Counseling status extraction changed"
            n = n + 1
            id = ImportPDFPreview("20", "SSN"): Set saved = SavedFields(id)
            If saved("CounselingDate") <> status Or saved("FormKind") <> kind Then Err.Raise vbObjectError + 4752, , "Counseling status import changed"
            n = n + 1
        Next
    Next
    TestPDFCounselingStatuses = "PASS: " & n & " PDF counseling status extraction/save assertions": Exit Function
Failed:
    TestPDFCounselingStatuses = "FAIL after " & n & ": " & Err.Number & " / " & Err.Description
End Function
