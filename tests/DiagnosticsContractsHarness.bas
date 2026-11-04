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
Private Sub CheckWizardViewContracts()
    Dim item As Variant, view As String, c As Access.Control, original As String, binding As String
    For Each item In Array(Array("frmUXWorkspaceImport", "txtProgress", "cmdApply"), Array("frmPDFImport", "uxState", "cmdImport"), Array("frmNAVFITExchange", "txtStatus", "cmdExportGroup"), Array("frmPDFOutput", "lblStatus", "cmdClose"), Array("frmUXComparison", "lblStatus", "cmdDiscardCells"), Array("frmMyWork", "uxWorkStatus", "cmdOpenWork"), Array("frmUXNativeGroups", "lblStatus", "cmdApply"), Array("frmUXMoveGroup", "lblStatus", "cmdApply"), Array("frmUXDirectoryMerge", "txtPreview", "cmdApply"))
        view = CStr(item(0)): DoCmd.OpenForm view, acDesign
        Check UXAssertView(view), "Generated wizard contract rejected: " & view
        Set c = Forms(view).Controls(CStr(item(1))): original = c.Name: c.Name = "SyntheticMissingStatus"
        Check ContractRejected(view, 9400), "Missing required progress/status accepted: " & view
        c.Name = original
        Set c = Forms(view).Controls(CStr(item(2))): binding = c.OnClick: c.OnClick = "=CloseEditor()"
        Check ContractRejected(view, 9401), "Wrong wizard action accepted: " & view
        c.OnClick = binding
        Check UXAssertView(view), "Restored wizard contract rejected: " & view
        DoCmd.Close acForm, view, acSaveNo
    Next
End Sub
Private Sub CheckDynamicViewContracts()
    Dim item As Variant, view As String, c As Access.Control, original As String
    For Each item In Array(Array("frmLibrary", "uxNav5"), Array("frmContextDuties", "txtDutyDescription"), Array("frmUXEditorFields", "errOfficialName"), Array("frmUXSeniorEditor", "rsReportingSeniorEmail"), Array("frmUXMemberEditor", "ruadPage4"), Array("frmUXSharedFields", "valueUnitUIC"), Array("frmUXRSCATable", "countO10"), Array("frmUXMemberHistory", "sort_lstCounseling_4"), Array("frmServiceDetails", "txtReadinessDetail"), Array("frmUXPriorEditor", "lblTrait39"), Array("frmEditor", "uxSection9"))
        view = CStr(item(0)): DoCmd.OpenForm view, acDesign
        Check UXAssertView(view), "Dynamic view rejected: " & view
        Set c = Forms(view).Controls(CStr(item(1))): original = c.Name: c.Name = "SyntheticDynamicMissing"
        Check ContractRejected(view, 9400), "Missing computed control accepted: " & view & "." & original
        c.Name = original
        Check UXAssertView(view), "Restored computed control rejected: " & view
        DoCmd.Close acForm, view, acSaveNo
    Next
End Sub
Private Function SurfaceAssignmentError(ByVal c As Access.Control, ByVal value As Variant, ByRef message As String) As Long
    On Error GoTo Failed
    c.Value = value
    Exit Function
Failed:
    SurfaceAssignmentError = Err.Number: message = Err.Description
End Function
Private Sub CheckSurfaceFailure(ByVal reportID As Long)
    Dim f As Access.Form, surface As Access.Form, c As Access.Control, key As String, binding As String
    Dim expected As Long, message As String, refreshFailure As Object, restoreFailure As Object, refreshed As Boolean, editorRefreshed As Boolean, pageRefreshed As Boolean, revision As Long, pending As Boolean, savedValue As Variant, pageFailure As Object
    Set f = Forms!frmEditor: Set surface = f!uxFieldSurface.Form
    key = "OfficialName": Set c = surface.Controls(UXFieldControl(key))
    revision = CurrentRevision(reportID): savedValue = SavedFields(reportID)(key)
    Check UXSurfaceSetValue(key, "SYNTHETIC pending surface"), "Could not stage surface pending value"
    pending = EditorHasPending()
    Check pending, "Surface fixture has no pending value"
    f!cmdClose.SetFocus
    binding = c.ControlSource: c.ControlSource = "=1"
    'A calculated field cannot accept the buffer value. Observe the native cause.
    expected = SurfaceAssignmentError(c, f.Controls(c.Name).Value, message)
    refreshed = UXSurfaceRefresh(): Set refreshFailure = SwiftLastFailureContext()
    Call UXSurfaceRestoreField(key): Set restoreFailure = SwiftLastFailureContext()
    editorRefreshed = UXRefreshEditor()
    pageRefreshed = UXFixedSelect(0): Set pageFailure = SwiftLastFailureContext()
    c.ControlSource = binding
    Check expected <> 0, "Synthetic surface assignment did not fail"
    Check Not refreshed And refreshFailure("Procedure") = "UXSurfaceRefresh", "Surface refresh suppressed its primary diagnostic"
    Check refreshFailure("Number") = expected And refreshFailure("Description") = message, "Surface refresh changed native error number or description"
    Check refreshFailure("Stage") = "Refresh field OfficialName" And refreshFailure("Transaction") = "Not started", "Surface refresh lost field stage or transaction outcome"
    Check restoreFailure("Procedure") = "UXSurfaceRestoreField" And restoreFailure("Number") = expected, "Field restore suppressed its primary diagnostic"
    Check restoreFailure("Stage") = "Restore field OfficialName", "Field restore lost its stage"
    Check Not editorRefreshed, "Editor ignored failed surface refresh"
    Check Not pageRefreshed And pageFailure("Procedure") = "UXSurfaceRefresh", "Page ignored or replaced failed surface diagnostic"
    Check f.Painting And surface.Painting, "Surface failure left page painting suspended"
    Check InStr(f!uxError.Caption, "OfficialName") > 0 And InStr(f!uxError.ControlTipText, "UXSurfaceRefresh") > 0, "Surface failure not visible"
    Check CurrentRevision(reportID) = revision And EditorHasPending() = pending, "Surface failure changed saved or pending state"
    Check f.Controls(c.Name).Value = "SYNTHETIC pending surface" And SameFieldValue(SavedFields(reportID)(key), savedValue), "Surface failure lost pending value or saved snapshot"
    Check UXSurfaceRefresh() And UXRefreshEditor(), "Surface retry stayed busy after failure"
    Check SameFieldValue(c.Value, f.Controls(c.Name).Value), "Surface retry did not restore displayed buffer value"
    Check UXUndoPending(), "Could not close disposable pending fixture"
End Sub
'The save commits before the display refresh, so a refresh failure is reported as saved.
Private Sub CheckSaveRefreshFailure(ByVal reportID As Long)
    Dim f As Access.Form, surface As Access.Form, c As Access.Control, key As String, binding As String, revision As Long, saved As Boolean, context As Object
    Set f = Forms!frmEditor: Set surface = f!uxFieldSurface.Form
    key = "OfficialName": Set c = surface.Controls(UXFieldControl(key))
    revision = CurrentRevision(reportID)
    Check UXSurfaceSetValue(key, "SURFACE, SAVED"), "Could not stage surface save value"
    f!cmdClose.SetFocus
    binding = c.ControlSource: c.ControlSource = "=1"
    saved = UXSaveEditor(): Set context = SwiftLastFailureContext()
    c.ControlSource = binding
    Check saved And CurrentRevision(reportID) <> revision, "Save failed when only the display refresh failed"
    Check SavedFields(reportID)(key) = "SURFACE, SAVED", "Saved value lost after display refresh failure"
    Check context("Procedure") = "UXSurfaceRefresh" And context("Transaction") = "Committed", "Post-save refresh failure lost committed outcome"
    Check InStr(f!uxError.Caption, "Saved; display refresh failed") = 1 And Not EditorHasPending(), "Post-save refresh failure claimed unsaved edits"
    Check UXSurfaceRefresh() And UXRefreshEditor(), "Display retry failed after post-save refresh failure"
End Sub
Private Sub CheckPageFailure(ByVal reportID As Long)
    Dim f As Access.Form, surface As Access.Form, context As Object, revision As Long, actor As Long
    Set f = Forms!frmEditor: Set surface = f!uxFieldSurface.Form
    revision = CurrentRevision(reportID): actor = MajorAuthor()
    'A missing saved row fails after painting is suspended; no source injection.
    f!txtReportID = 2147483647
    Check Not UXFixedSelect(0), "Missing report page falsely succeeded"
    Set context = SwiftLastFailureContext()
    Check context("Number") = 3021 And context("Stage") = "Read report state", "Page lost primary error or stage"
    Check InStr(f!uxError.Caption, "Read report state") > 0 And InStr(f!uxError.ControlTipText, "3021") > 0, "Page diagnostic not visible"
    Check f.Painting And surface.Painting, "Failed page left painting suspended"
    f!txtReportID = reportID
    Check CurrentRevision(reportID) = revision And Not EditorHasPending(), "Page failure changed report or pending state"
    Check UXFixedSelect(0), "Page retry failed after restoring context"
    'The review callback returns False rather than raising into its caller.
    ResetActiveContributor
    Check Not UXHeaderReviewState(), "Review refresh without contributor falsely succeeded"
    Set context = SwiftLastFailureContext()
    Check context("Number") = vbObjectError + 8900 And context("Procedure") = "UXHeaderReviewState", "Review lost primary error"
    Check InStr(f!uxError.Caption, "Review actions unavailable") = 1, "Review failure silently suppressed"
    Check Not UXFixedSelect(0), "Page ignored failed review callback"
    Check f.Painting And surface.Painting, "Failed review left painting suspended"
    BindActiveContributor actor
    Check UXFixedSelect(0) And UXHeaderReviewState(), "Review retry failed after contributor restore"
    Check CurrentRevision(reportID) = revision And Not EditorHasPending(), "Review failure changed saved or pending data"
End Sub
'A disposable dialog event expires attribution only after the signing commit.
Public Function DiagnosticsExpireContributor() As Boolean
    ResetActiveContributor
    DiagnosticsExpireContributor = True
End Function
Private Sub CheckCommittedRefresh(ByVal reportID As Long)
    Dim actor As Long, revision As Long, context As Object, f As Access.Form
    actor = MajorAuthor(): revision = CurrentRevision(reportID): Set f = Forms!frmEditor
    Check UXSignoffOpen(reportID), "Synthetic signing dialog did not open"
    Forms!frmUXSignoff!txtSignedOn = Date
    Forms!frmUXSignoff!chkDispose = False
    Forms!frmUXSignoff.OnClose = "=DiagnosticsExpireContributor()"
    Check UXSignoffSave(), "Committed signing falsely reported failure"
    Check Not HasActiveContributor(), "Synthetic postcommit event did not run"
    Set context = SwiftLastFailureContext()
    Check context("Number") = vbObjectError + 8900 And context("Transaction") = "Committed", "Nested refresh lost primary error or committed outcome"
    Check InStr(f!uxError.Caption, "Saved; editor refresh failed") = 1 And InStr(f!uxError.ControlTipText, "Committed") > 0, "Committed refresh failure not visible"
    BindActiveContributor actor
    Check ReportIsSigned(reportID) And CurrentRevision(reportID) = revision, "Refresh failure changed signed saved revision"
    Check Not EditorHasPending() And Not f!cmdSave.Enabled, "Refresh failure unlocked signed report or added pending edits"
    Check UXFixedSelect(0), "Signed page retry failed"
End Sub
Private Sub CheckContributorFailures()
    Dim actor As Long, item As Variant, f As Access.Form, result As Boolean, status As String
    actor = MajorAuthor()
    For Each item In Array("frmHome", "frmMyWork")
        DoCmd.OpenForm CStr(item): Set f = Forms(CStr(item))
        If CStr(item) = "frmHome" Then status = "uxContributorStatus" Else status = "uxWorkStatus"
        f!uxContributor = 0
        If CStr(item) = "frmHome" Then result = UXOverviewContributor() Else result = UXWorkContributor()
        Check Not result And f.Controls(status).Caption = "Select an existing contributor.", "Contributor rejection lost original message: " & CStr(item)
        Check Nz(f!uxContributor, 0) = actor And MajorAuthor() = actor, "Rejected switch changed active contributor"
        ResetActiveContributor
        f!uxContributor = 0
        If CStr(item) = "frmHome" Then result = UXOverviewContributor() Else result = UXWorkContributor()
        Check Not result And f.Controls(status).Caption = "Select an existing contributor.", "Missing-session contributor rejection masked error: " & CStr(item)
        Check IsNull(f!uxContributor) And Not HasActiveContributor(), "Missing session displayed a confirmed contributor"
        BindActiveContributor actor
        DoCmd.Close acForm, CStr(item), acSaveNo
    Next
End Sub
Public Function TestDiagnosticsContracts() As String
    Dim context As Object, workspace As DAO.Workspace, db As DAO.Database, r As DAO.Recordset
    Dim saved As Long, before As Long, c As Access.Control, binding As String, operation As Long
    On Error GoTo Failed
    checks = 0
    Set context = SwiftFailure("TestModule", "TestProcedure", "Write", 12345, "Synthetic 123-45-6789 / 123456789 / 123 45 6789 / 1234567890", "Native source", "Active")
    Check context("Description") Like "Synthetic*", "Original description retained in owned memory"
    'Plain-VBA redaction (no VBScript); expected values are the former RegExp results.
    Check SwiftRedact("call 555 123 4567 now") = "call [identifier redacted] now" And SwiftRedact("1-2-3-4-5-6-7-8-9") = "[identifier redacted]", "Separated identifier not redacted"
    Check SwiftRedact("12345678") = "12345678" And SwiftRedact("123  456789") = "123  456789", "Short or broken digit runs redacted"
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
    CheckSurfaceFailure saved
    CheckSaveRefreshFailure saved
    CheckPageFailure saved
    CheckCommittedRefresh saved
    DoCmd.Close acForm, "frmEditor", acSaveNo
    CheckContributorFailures
    CheckWizardViewContracts
    CheckDynamicViewContracts
    TestDiagnosticsContracts = "PASS: " & checks & " primary/cleanup diagnostics, rollback/retry and missing/renamed/binding contracts"
    Exit Function
Failed:
    TestDiagnosticsContracts = "FAIL: " & Err.Number & " " & Err.Description
End Function
