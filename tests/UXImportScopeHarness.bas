Attribute VB_Name = "UXImportScopeHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6930, , message
    n = n + 1
End Sub
Private Function Receipt(ByVal id As Long, ByVal adapter As String) As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("ImportReceipt", dbOpenDynaset)
    r.AddNew: r!ReceiptGuid = NewIdentity(): r!SourceHash = String$(64, "0"): r!SourcePath = "Synthetic scope fixture"
    r!ExternalReportID = n + 1: r!ReportID = id: r!AuthorProfileID = MajorAuthor(): r!ImportedAt = Now(): r!AdapterVersion = adapter
    r!AppliedRevisionID = CurrentRevision(id): r!IsUndone = False: r.Update: r.Bookmark = r.LastModified: Receipt = r!ImportReceiptID: r.Close
End Function
Public Function TestUXImportScope() As String
    Dim id As Long, last As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC IMPORT SCOPE", "Tester"
    id = CommitDraft(0, "EVAL", "Scope fixture", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Check NAVFITImportScope(id) = "", "Manual report labelled imported"
    Call Receipt(id, "comments-v1/local-CStr")
    Check NAVFITImportScope(id) = "Narrative only", "Legacy narrative not classified"
    OpenDraftChecklist id
    Check InStr(Forms!frmChecklist!lstFindings.RowSource, "PARTIAL_IMPORT") > 0, "Legacy partial finding missing"
    DoCmd.Close acForm, "frmChecklist", acSaveNo
    last = Receipt(id, "navfit-full-v1/local-CStr")
    Check NAVFITImportScope(id) = "Full report", "Full mapping classified as narrative"
    OpenDraftChecklist id
    Check InStr(Forms!frmChecklist!lstFindings.RowSource, "PARTIAL_IMPORT") = 0, "False partial finding for full mapping"
    DoCmd.Close acForm, "frmChecklist", acSaveNo
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Check InStr(Forms!frmEditor!lblImportOrigin.Caption, "Full NAVFIT report") > 0, "Editor uses stale scope label"
    UXLeaveEditor "Discard"
    CurrentDb.Execute "UPDATE ImportReceipt SET IsUndone=True WHERE ImportReceiptID=" & last, dbFailOnError
    Check NAVFITImportScope(id) = "Narrative only", "Undone full receipt still controls scope"
    DoCmd.OpenForm "frmUXImportChoose"
    Check InStr(Forms!frmUXImportChoose!cmdNAVFIT.Caption, "groups and reports") > 0, "Chooser advertises narrative-only"
    TestUXImportScope = "PASS: " & n & " import scope assertions": Exit Function
Failed:
    TestUXImportScope = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
