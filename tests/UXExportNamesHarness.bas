Attribute VB_Name = "UXExportNamesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6913, , message
    n = n + 1
End Sub
Public Function TestUXExportNames() As String
    Dim id As Long, other As Long, g As Long, fields As Object, fs As Object, folder As String, first As String, second As String, file As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC EXPORT NAMES", "Tester"
    id = CommitDraft(0, "EVAL", "Names", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "RankRate", "IT2"
    CommitSavedFields id, CurrentRevision(id), fields, "Worn rate"
    g = CreateSummaryGroup("Export names", "EVAL", "E5"): AddGroupMember g, id, "E5"
    Check UXExportName(0, id) = "Swift_NAVFIT_Export_IT2_2026-09-30.accdb", "Single report filename wrong"
    Check UXExportName(g, 0) = "Swift_NAVFIT_Export_E5_2026-09-30.accdb", "Group filename wrong"
    other = CommitDraft(0, "EVAL", "Mixed", "Other synthetic member", #1/1/2026#, #8/31/2026#, "", "Test", "")
    AddGroupMember g, other, "E5"
    Check UXExportName(g, 0) = "Swift_NAVFIT_Export_E5_MIXED.accdb", "Mixed ending dates mislabeled"
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.BuildPath(CurrentProject.Path, "export-name-test-" & NewIdentity()): fs.CreateFolder folder
    first = UXExportDestination(folder, 0, id)
    Check fs.GetFileName(first) = UXExportName(0, id), "Directory selection did not generate filename"
    Set file = fs.CreateTextFile(first, False): file.Write "synthetic existing file": file.Close
    second = UXExportDestination(folder, 0, id)
    Check second <> first And InStr(second, "_1.accdb") > 0, "Existing export would be overwritten"
    fs.DeleteFile first
    Set file = fs.CreateTextFile(first & ".swift-baseline.accdb", False): file.Close
    Check UXExportDestination(folder, 0, id) = second, "Orphan baseline would be overwritten"
    Check UXExportDestination(first, 0, id) = second, "Subsequent export did not retain directory"
    fs.DeleteFile first & ".swift-baseline.accdb": RmDir folder
    DoCmd.OpenForm "frmNAVFITExchange"
    ' The path box holds a directory (typed or chosen); UXExportDestination always supplies the filename.
    Check Forms!frmNAVFITExchange!lbltxtExportPath.Caption = "Export directory", "Export path not labeled as a directory"
    Check Forms!frmNAVFITExchange!txtExportPath.AfterUpdate = "=UXPreviewNAVFITGroup()", "Typed export directory does not refresh the preview"
    Check Forms!frmNAVFITExchange!cmdtxtExportPath.Caption = "Choose folder", "Export chooser still labeled file selection"
    Check Forms!frmNAVFITExchange!cmdtxtExportPath.OnClick = "=UXChooseNAVFITGroupFolder()", "Export chooser does not pick a folder"
    TestUXExportNames = "PASS: " & n & " automatic export filename assertions": Exit Function
Failed:
    TestUXExportNames = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
