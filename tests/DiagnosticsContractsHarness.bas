Attribute VB_Name = "DiagnosticsContractsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 9500, "DiagnosticsContractsHarness", label
    checks = checks + 1
End Sub
Private Function RaisedPrimary(ByVal context As Object) As Boolean
    On Error GoTo Failed
    SwiftRaise context
    Exit Function
Failed:
    RaisedPrimary = Err.Number = CLng(context("Number")) And InStr(Err.Description, context("Identity")) > 0
End Function
Private Function ContractRejected(ByVal view As String, ByVal code As Long) As Boolean
    On Error GoTo Failed
    Call UXAssertView(view)
    Exit Function
Failed:
    ContractRejected = Err.Number = vbObjectError + code And InStr(Err.Description, view) > 0
End Function
Private Function FailedSave() As Boolean
    On Error GoTo Failed
    Call CommitDraft(0, "EVAL", "Synthetic", "Synthetic Member", Null, Null, "", "Test", "Synthetic narrative")
    Exit Function
Failed:
    FailedSave = InStr(SwiftLastFailure(), "CommitDraft") > 0 And InStr(SwiftLastFailure(), "Rolled back") > 0
End Function
Public Function TestDiagnosticsContracts() As String
    Dim context As Object, workspace As DAO.Workspace, db As DAO.Database, r As DAO.Recordset
    Dim saved As Long, before As Long, c As Access.Control, binding As String, operation As Long
    On Error GoTo Failed
    checks = 0
    Set context = SwiftFailure("TestModule", "TestProcedure", "Write", 12345, "Synthetic 123-45-6789 / 123456789 / 123 45 6789 / 1234567890", "Native source", "Active")
    Check context("Description") Like "Synthetic*", "Original description retained in owned memory"
    Check InStr(SwiftFailureText(context), "123-45-6789") = 0 And InStr(SwiftFailureText(context), "123456789") = 0 And InStr(SwiftFailureText(context), "123 45 6789") = 0, "Identifiers redacted from diagnostics"
    Set workspace = DBEngine.CreateWorkspace("SyntheticDiagnostics", "admin", "", dbUseJet)
    SwiftRollback workspace, context
    Check context("Number") = 12345 And context("Source") = "Native source", "Cleanup preserves primary number/source"
    Check InStr(context("Transaction"), "Rollback failed") > 0 And InStr(context("Cleanup"), "rollback") > 0, "Rollback failure reported separately"
    Check RaisedPrimary(context), "Reraise preserves primary identity"
    workspace.Close
    Set workspace = DBEngine.Workspaces(0): Set db = CurrentDb
    workspace.BeginTrans
    db.Execute "UPDATE AppMetadata SET CompatibilityStatus='Synthetic rollback'", dbFailOnError
    SwiftRollback workspace, context
    Check context("Transaction") = "Rolled back", "Successful rollback has explicit outcome"
    Set r = db.OpenRecordset("SELECT * FROM AppMetadata", dbOpenSnapshot): r.Close
    SwiftClose r, context, "already closed reader"
    Check context("Number") = 12345, "Reader cleanup cannot replace primary"
    Call ConfigureWorkspace("SYNTHETIC DIAGNOSTICS", "Synthetic Contributor")
    operation = BeginOperation("Synthetic failure", "Synthetic test")
    db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = "<> 'Failed'"
    SwiftFinishFailed operation, context
    Check context("Number") = 12345 And InStr(context("Cleanup"), "record failed operation") > 0, "Journal cleanup failure preserves primary"
    Check DLookup("OperationState", "OperationJournal", "OperationID=" & operation) = "Started", "Failed diagnostic write rolls back"
    db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = ""
    db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'OfficialComments'"
    before = DCount("*", "Report")
    Check FailedSave(), "Native failed write retains diagnostic and rollback outcome"
    Check DCount("*", "Report") = before And DCount("*", "ReportRevision") = 0, "Late save failure rolls back domain/revision rows"
    db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    saved = CommitDraft(0, "EVAL", "Synthetic", "Synthetic Member", Null, Null, "", "Retry", "Synthetic narrative")
    Check saved > 0 And DCount("*", "ReportRevision") = 1, "Retry succeeds after correction"
    DoCmd.OpenForm "frmEditor", acDesign
    Check UXAssertView(UX_EDITOR_VIEW), "Editor contract accepts generated view"
    Set c = Forms!frmEditor!cmdSave: c.Name = "SyntheticRenamedSave"
    Check ContractRejected(UX_EDITOR_VIEW, 9400), "Renamed required control detected"
    c.Name = "cmdSave": binding = c.OnClick: c.OnClick = "=CloseEditor()"
    Check ContractRejected(UX_EDITOR_VIEW, 9401), "Existing but incorrect callback detected"
    c.OnClick = binding
    DoCmd.Close acForm, "frmEditor", acSaveNo
    DoCmd.OpenForm "frmUXTraitRows", acDesign
    Check UXAssertView(UX_TRAIT_ROWS_VIEW), "Dynamic trait fields have queue contracts"
    DoCmd.Close acForm, "frmUXTraitRows", acSaveNo
    DoCmd.OpenForm "frmEditor", , , , , , CStr(saved)
    Call UXShowSaveError("Synthetic display failure" & vbCrLf & "Diagnostic: synthetic; transaction: Committed", True)
    Check InStr(Forms!frmEditor!uxError.Caption, "Saved;") = 1 And InStr(Forms!frmEditor!uxDirty.Caption, "Saved") = 1, "Committed save is not reported as unsaved"
    Check InStr(Forms!frmEditor!uxError.ControlTipText, "Committed") > 0, "Detailed outcome available without resizing status"
    Call UXShowSaveError("Synthetic write failure")
    Check InStr(Forms!frmEditor!uxError.Caption, "Not saved:") = 1, "Failed write keeps explicit unsaved feedback"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    TestDiagnosticsContracts = "PASS: " & checks & " primary/cleanup diagnostics, rollback/retry and missing/renamed/binding contracts"
    Exit Function
Failed:
    TestDiagnosticsContracts = "FAIL: " & Err.Number & " " & Err.Description
End Function
