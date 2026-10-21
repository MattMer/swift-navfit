Attribute VB_Name = "LocalPDFHarness"
Option Explicit
Public Function CheckLocalPDF(ByVal path As String, ByVal kind As String, ByVal templateHash As String, ByVal dataHash As String, ByVal rejection As String) As String
    Dim packets As Object, preview As Object, bytes() As Byte, before As Long, stamp As String, status As String
    On Error GoTo Failed
    before = DCount("*", "Report")
    Set preview = PreviewStructuredPDF(path)
    If Len(rejection) > 0 Then Err.Raise 5, , "Accepted a rejected fixture."
    If preview("FormKind") <> kind Or preview("TemplateHash") <> templateHash Then Err.Raise 5, , "Template identity differs from independent oracle."
    Set packets = PDFLocalPackets(path): bytes = packets("datasets")
    If PDFBytesHash(bytes) <> dataHash Then Err.Raise 5, , "Data bytes differ from independent oracle."
    stamp = PDFImportStamp(): status = PDFExtractionStatus(path)
    If Left$(status, 5) <> "PASS:" Or stamp <> PDFImportStamp() Or PDFImportKind() <> kind Then Err.Raise 5, , "Diagnostic changed preview."
    If DCount("*", "Report") <> before Then Err.Raise 5, , "Preview wrote report data."
    CheckLocalPDF = "PASS: exact packets, identity, diagnostic and read-only preview"
    Call PDFClearPreview: Exit Function
Failed:
    status = Err.Description
    If Len(rejection) > 0 And InStr(1, status, rejection, vbTextCompare) > 0 And status <> "Accepted a rejected fixture." Then
        CheckLocalPDF = "PASS: rejected - " & status
    Else
        CheckLocalPDF = "FAIL: " & status
    End If
    Call PDFClearPreview
End Function
