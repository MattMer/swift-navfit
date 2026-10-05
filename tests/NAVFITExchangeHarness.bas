Attribute VB_Name = "NAVFITExchangeHarness"
Option Compare Binary
Option Explicit
Private assertions As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    assertions = assertions + 1
    If Not value Then Err.Raise vbObjectError + 5400, , message
End Sub
Private Function ApplyRejected() As String
    On Error GoTo Failed
    Call NAVFITApplyReturned
    ApplyRejected = "ACCEPTED": Exit Function
Failed:
    ApplyRejected = Err.Description
End Function
Public Function TestNAVFITExchange(ByVal directory As String, ByVal template As String) As String
    Dim r As DAO.Recordset, source As DAO.Database, target As DAO.Database, a As DAO.Recordset, b As DAO.Recordset, id As Long, output As String, incoming As String, rows As Collection, row As Variant, key As Variant, changes As Object, count As Long, v As Variant
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM Report ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        id = r!ReportID: output = directory & "\export-" & id & ".accdb"
        Check NAVFITExportReport(id, output, template) = output, "Export failed"
        Check Len(Dir$(output & ".swift-baseline.accdb")) > 0, "Durable baseline missing"
        Set source = DBEngine.OpenDatabase(output, False, True)
        Set a = source.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
        Check Not a.EOF, "Export has no report"
        Check NAVFITKind(a!ReportType) = SavedFields(id)("FormKind"), "Export type changed"
        a.MoveNext: Check a.EOF, "Export includes unselected members"
        a.Close: source.Close
        Set source = DBEngine.OpenDatabase(Nz(DLookup("SourcePath", "ImportReceipt", "ReportID=" & id), ""), False, True)
        Set target = DBEngine.OpenDatabase(output, False, True)
        Set a = source.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & DLookup("ExternalReportID", "ImportReceipt", "ReportID=" & id), dbOpenSnapshot)
        Set b = target.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
        For Each key In a.Fields
            Check SameFieldValue(key.Value, b.Fields(key.Name).Value), "Export native field changed: " & key.Name
        Next
        a.Close: b.Close: source.Close: target.Close
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Export saved an extra revision"
        r.MoveNext
    Loop
    r.Close
    output = directory & "\export-1.accdb": incoming = directory & "\returned.accdb"
    FileCopy output, incoming
    Set target = DBEngine.OpenDatabase(incoming)
    target.Execute "UPDATE Reports SET Comments='NAVFIT returned comments', PROF=5 WHERE ReportType Is Not Null", dbFailOnError
    target.Close
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Swift newer comments"
    Call CommitSavedFields(1, CurrentRevision(1), changes, "Synthetic local edit after export")
    Set rows = NAVFITPreviewReturned(incoming, output & ".swift-baseline.accdb")
    count = 0
    For Each row In rows
        If row = "OfficialComments|Conflict" Then count = count + 1
    Next
    Check count = 1, "Distinct edits not shown as conflict"
    Check InStr(ApplyRejected(), "Resolve every") > 0, "Unresolved conflict applied"
    Check SavedFields(1)("OfficialComments") = "Swift newer comments", "Conflict rejection lost local value"
    NAVFITResolveReturned "OfficialComments", "Keep current"
    Check NAVFITApplyReturned() = 1, "Resolved return failed"
    Check SavedFields(1)("OfficialComments") = "Swift newer comments", "Keep current ignored"
    Check SavedFields(1)("Trait33") = "5.0", "Nonconflicting trait did not merge"
    Check DCount("*", "ImportReceipt", "ReportID=1") = 2, "Returned provenance missing"
    Check InStr(PreviewRejected(incoming, output & ".swift-baseline.accdb"), "newer native import") > 0, "Superseded native baseline reused"
    TestNAVFITExchange = "PASS: " & assertions & " NAVFIT exchange assertions": Exit Function
Failed:
    TestNAVFITExchange = "FAIL: " & Err.Number & " " & Err.Description
End Function
Private Function PreviewRejected(ByVal path As String, ByVal baseline As String) As String
    On Error GoTo Failed
    Call NAVFITPreviewReturned(path, baseline)
    PreviewRejected = "ACCEPTED": Exit Function
Failed:
    PreviewRejected = Err.Description
End Function
