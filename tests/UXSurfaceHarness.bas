Attribute VB_Name = "UXSurfaceHarness"
Option Compare Binary
Option Explicit
Private count As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 5200, , message
    count = count + 1
End Sub
Private Function SurfaceFits(ByVal f As Access.Form) As Boolean
    Dim c As Access.Control
    For Each c In f.Controls
        If c.Visible Then
            If c.Left < 0 Or c.Left + c.Width > f.Width Then Exit Function
        End If
    Next
    SurfaceFits = True
End Function
Public Function TestUXSurface() As String
    Dim id As Long, f As Access.Form, s As Access.Form, revision As Long, changes As Object, key As Variant, kind As Variant, section As Long, stage As String, snapshot As Long, groupID As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC SURFACE TEST", "Surface tester"
    id = CommitDraft(0, "EVAL", "Surface title", "Surface member", #1/1/2026#, #9/30/2026#, "Saved notes", "Test", "Saved comments")
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "MemberDoDID", "1234567890": changes.Add "ReportingSeniorDoDID", "0987654321"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Legacy identifiers retained")
    stage = "editor load": DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
    Check UXSurfaceAvailable(), "Native section surface unavailable"
    DoCmd.Maximize
    Check f!uxFieldSurface.Left = 5900 And f!uxFieldSurface.Width = f.InsideWidth - 6250, "Maximized field panel moved away from section content"
    Check f!cmdSave.Height = 500 And f!cmdClose.Height = 500, "Maximized action buttons stretched"
    Check f!uxFieldSurface.Top + f!uxFieldSurface.Height <= f!lblSaved.Top, "Field panel overlaps save area"
    Check f!uxFieldSurface.HorizontalAnchor = acHorizontalAnchorLeft And f!uxFieldSurface.VerticalAnchor = acVerticalAnchorTop, "Field panel uses moving anchors"
    Check f!cmdSave.VerticalAnchor = acVerticalAnchorTop And f!cmdClose.VerticalAnchor = acVerticalAnchorTop, "Action buttons use stretching anchors"
    Check InStr(f!uxContext.Caption, "Surface tester") > 0, "Workspace context missing contributor"
    Check f!cmdSave.Top = f.InsideHeight - 1300, "Footer did not follow viewport height"
    Check s.Width = f!uxFieldSurface.Width - 300, "Child fields did not follow panel width"
    Check f!uxGroup.Width = 4000, "Summary group selector stayed narrow"
    If s.Width >= 14500 Then Check s!txtMemberMiddleInitial.Top = s!txtMemberFirstName.Top, "Wide editor did not use three field columns"
    DoCmd.Restore
    Check f!cmdSave.Height = 500 And f!uxFieldSurface.Left = 5900, "Restoring window displaced controls"

    Check s!txtOfficialName.Visible And Not f!txtOfficialName.Visible, "Editor did not use native section surface"
    Check Not f!uxNext.Visible And Not f!uxPrevious.Visible, "Legacy field pagination still visible"
    Check UXResizeEditor(25000, 13500), "Wide viewport layout failed"
    Check f!uxFieldSurface.Width = 18750 And f!cmdSave.Top = 12200, "Wide viewport not filled"
    Check s!txtMemberMiddleInitial.Top = s!txtMemberFirstName.Top, "Wide layout did not switch to three columns"
    Check UXResizeEditor(13800, 8650), "Laptop viewport layout failed"
    Check f!uxFieldSurface.Width = 7550 And s.Width = 7250, "Laptop fields did not shrink to panel"
    Check s!txtMemberMiddleInitial.Top > s!txtMemberFirstName.Top, "Laptop layout did not restore two columns"
    Call UXResizeEditor()

    revision = CurrentRevision(id)
    stage = "section switching": f!txtMemberFirstName = "Separate first name"
    For section = 0 To 7
        Check UXSelectSection(section), "Section failed to open"
        Check SurfaceFits(s), "Section controls exceeded field surface width"
        Check f!txtMemberFirstName = "Separate first name", "Section switch lost pending name parts"
        Check CurrentRevision(id) = revision, "Section switch saved a revision"
    Next
    Call UXSelectSection(0)
    Check s!txtMemberFirstName = "Separate first name", "Visible field did not reflect pending buffer"
    Check IsNull(f!txtOfficialName), "Changing name parts overwrote official full name"
    stage = "active text": f!uxFieldSurface.SetFocus: s!txtOfficialName.SetFocus: s!txtOfficialName.Text = "SURFACE TESTER"
    Check EditorHasPending(), "Active surface text not detected as pending"
    Call UXResizeEditor()
    Check f!txtOfficialName = "SURFACE TESTER" And EditorHasPending(), "Resize lost pending active text"
    Check CurrentRevision(id) = revision, "Resize committed a revision"
    Check CheckpointEditor(), "Active surface text checkpoint failed"
    Check DLookup("OfficialName", "RecoveryDraft", "ReportID=" & id) = "SURFACE TESTER", "Checkpoint lost active surface text"
    Check CurrentRevision(id) = revision, "Checkpoint saved report revision"
    Call UXSelectSection(2)
    Check f!txtOfficialName = "SURFACE TESTER", "Leaving section lost active text"
    stage = "checkboxes": Call UXSelectSection(0): s!txtOccasionPeriodic = True
    Check UXSurfaceEdited("OccasionPeriodic"), "Checkbox bridge failed"
    Check f!txtOccasionPeriodic = "Yes", "Checkbox did not map to saved flag"
    s!cboObservationStatus = False: Call UXSurfaceEdited("ObservationStatus")
    Check f!cboObservationStatus = "Observed", "Unchecked NOB did not map to observed"
    s!cboObservationStatus = Null: Call UXSurfaceEdited("ObservationStatus")
    Check IsNull(f!cboObservationStatus), "Indeterminate NOB became an explicit value"
    stage = "inline validation": f!txtMemberSSN = "123": UXSurfaceRefresh
    Check Not UXSaveEditor(), "Invalid SSN saved"
    Check CurrentRevision(id) = revision And f!txtMemberSSN = "123", "Failed save lost pending values or changed revision"
    Check f!uxErrorList.Visible And InStr(s!errMemberSSN.Caption, "9 digits") > 0, "Inline and summary errors missing"
    f!uxErrorList = "MemberSSN": Check UXSurfaceErrorSelected(), "Error summary did not open field"
    Check s!txtMemberSSN.Visible, "Error target remains hidden"
    f!cmdClose.SetFocus: f!txtMemberSSN = "000000099": UXSurfaceRefresh
    stage = "save": Check UXSaveEditor(), "Valid surface draft save failed"
    Check CurrentRevision(id) <> revision, "Save did not append revision"
    Check Not EditorHasPending(), "Saved surface remains dirty"
    Check DCount("*", "RevisionValue", "RevisionID=" & CurrentRevision(id)) = 78, "Surface save lost snapshot fields"
    Check DLookup("MemberDoDID", "Report", "ReportID=" & id) = "1234567890", "Hidden member identifier lost"
    Check DLookup("ReportingSeniorDoDID", "Report", "ReportID=" & id) = "0987654321", "Hidden reporting senior identifier lost"
    Call UXSelectSection(5): Check s!txtOfficialComments.Visible And s!txtOfficialComments.Height >= 2000, "Narrative lacks spacious editor"
    Call UXSelectSection(4): Check s!uxTraitGuidance.Visible And s!uxTraitGuidance.Locked, "Trait guidance unavailable or editable"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    stage = "captions": For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        Check InStr(UXFieldCaption("OfficialComments", CStr(kind)), IIf(kind = "EVAL", "43.", IIf(kind = "CHIEFEVAL", "40.", "41."))) = 1, "Comments block caption incorrect"
        Check UXFieldCaption("Trait33", CStr(kind)) Like "33. *", "Trait block caption missing"
    Next
    stage = "dashboard": DoCmd.OpenForm "frmHome": Check UXDashboardRefresh(), "Dashboard failed to load"
    Check Forms!frmHome!uxRecentReports.ListCount >= 2, "Recent reports list empty"
    Check CLng(Forms!frmHome!uxReportCount.Caption) = 1, "Dashboard active report count incorrect"
    stage = "filters": DoCmd.OpenForm "frmLibrary"
    Forms!frmLibrary!uxTypeFilter = "FITREP": RefreshLibrary
    Check Forms!frmLibrary!lstReports.ListCount <= 1, "Type filter included another report type"
    Check UXClearReportFilters(), "Clear filters failed"
    Forms!frmLibrary!lstReports = id: RefreshLibrary
    Check Forms!frmLibrary!lstReports = id, "Refresh lost selected report"
    stage = "group workspace": groupID = CreateSummaryGroup("Surface group", "EVAL", "E5")
    AddGroupMember groupID, id, "E5"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = groupID
    Check UXOpenGroup(), "Group list did not open focused workspace"
    Check Not Forms!frmMajorWorkbench!cboGroup.Visible, "Group workspace still asks users to select a second group"
    Check InStr(Forms!frmMajorWorkbench!uxGroupTitle.Caption, "Surface group") > 0, "Group title missing"
    Check UXWorkbenchMode("Review") And Forms!frmMajorWorkbench!uxOpenGroupReport.Visible, "Group review lacks report entry point"
    Check UXWorkbenchMode("Comparison") And Forms!frmMajorWorkbench!uxComparison.Visible, "Group comparison unavailable"
    stage = "my work": DoCmd.OpenForm "frmMyWork"
    Check Forms!frmMyWork!lstTasks.Visible And Not Forms!frmMyWork!cboContributor.Visible, "My Work default mixes contributors and tasks"
    Check UXMyWorkMode("Assign"), "Assignment workspace unavailable"
    Forms!frmMyWork!uxAssignReport = id: Check UXChooseAssignmentReport(), "Named report selection failed"
    Check Forms!frmMyWork!txtExpectedRevision = CurrentRevision(id), "Assignment did not capture saved revision"
    Check UXMyWorkMode("Contributors") And Forms!frmMyWork!cboContributor.Visible And Not Forms!frmMyWork!lstTasks.Visible, "Contributor workspace not focused"
    TestUXSurface = "PASS: " & count & " native surface, buffer, validation, captions and dashboard assertions": Exit Function
Failed:
    TestUXSurface = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
