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
' Fixed sheet geometry of every visible field control; resizing must leave it unchanged.
Private Function FieldGeometry(ByVal f As Access.Form) As String
    Dim c As Access.Control
    FieldGeometry = CStr(f.Width)
    For Each c In f.Controls
        If c.Visible Then FieldGeometry = FieldGeometry & "|" & c.Name & ":" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height
    Next
End Function
Public Function TestUXSurface() As String
    Dim id As Long, f As Access.Form, s As Access.Form, revision As Long, changes As Object, key As Variant, kind As Variant, section As Long, stage As String, snapshot As Long, groupID As Long, geometry As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC SURFACE TEST", "Surface tester"
    id = CommitDraft(0, "EVAL", "Surface title", "Surface member", #1/1/2026#, #9/30/2026#, "Saved notes", "Test", "Saved comments")
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "MemberDoDID", "1234567890": changes.Add "ReportingSeniorDoDID", "0987654321"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Legacy identifiers retained")
    stage = "editor load": DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
    Check UXSurfaceAvailable(), "Native section surface unavailable"
    geometry = FieldGeometry(s)
    Check f!uxFieldSurface.Left = 2850, "Field panel overlaps the navigation rail"
    DoCmd.Maximize
    Check FieldGeometry(s) = geometry And f!uxFieldSurface.Left = 2850, "Maximizing changed the field layout"
    Check f!cmdSave.Height = 450 And f!cmdClose.Height = 450, "Maximized action buttons stretched"
    Check f!uxFieldSurface.Top + f!uxFieldSurface.Height <= f!lblSaved.Top, "Field panel overlaps save area"
    Check f!uxFieldSurface.HorizontalAnchor = acHorizontalAnchorLeft And f!uxFieldSurface.VerticalAnchor = acVerticalAnchorTop, "Field panel uses moving anchors"
    Check f!cmdSave.VerticalAnchor = acVerticalAnchorTop And f!cmdClose.VerticalAnchor = acVerticalAnchorTop, "Action buttons use stretching anchors"
    Check InStr(f!uxContext.Caption, "Surface tester") > 0, "Workspace context missing contributor"
    Check f!cmdSave.Top = f!uxFieldSurface.Top + f!uxFieldSurface.Height + 490, "Footer did not follow the viewport"
    Check f!uxGroup.Width = 2850, "Summary group selector lost its fixed width"
    Check Not s!txtMemberMiddleInitial.Visible And Not s!txtMemberFirstName.Visible, "Redundant structured names visible"
    DoCmd.Restore
    Check FieldGeometry(s) = geometry And f!cmdSave.Height = 450 And f!uxFieldSurface.Left = 2850, "Restoring window displaced controls"

    Check s!txtOfficialName.Visible And Not f!txtOfficialName.Visible, "Editor did not use native section surface"
    Check Not f!uxNext.Visible And Not f!uxPrevious.Visible, "Legacy field pagination still visible"
    Check UXResizeEditor(25000, 13500), "Wide viewport layout failed"
    Check f!uxFieldSurface.Width = 21750 And f!cmdSave.Top = 13050, "Wide viewport not filled"
    Check FieldGeometry(s) = geometry, "Wide viewport changed the field layout"
    Check Not s!txtMemberFirstName.Visible, "Wide layout exposed redundant names"
    Check UXResizeEditor(13800, 8650), "Laptop viewport layout failed"
    Check f!uxFieldSurface.Width = 12600 And f.ScrollBars = 3, "Laptop viewport does not scroll the full sheet"
    Check FieldGeometry(s) = geometry, "Laptop viewport changed the field layout"
    Check Not s!txtMemberMiddleInitial.Visible, "Laptop exposed redundant middle initial"
    Check s!txtComponentActive.Top = s!txtComponentTAR.Top And s!txtComponentTAR.Top = s!txtComponentInactive.Top, "Component checkboxes did not share a compact row"
    Check s!labelComponentActive.Left > s!txtComponentActive.Left And Abs(s!labelComponentActive.Top - s!txtComponentActive.Top) <= 40, "Checkbox caption is not beside the checkbox"
    Call UXResizeEditor()

    revision = CurrentRevision(id)
    stage = "section switching": f!txtMemberFirstName = "Separate first name"
    For section = 0 To 7
        Check UXSelectSection(section), "Section failed to open"
        Check SurfaceFits(s), "Section controls exceeded field surface width"
        Check f!txtMemberFirstName = "Separate first name", "Section switch lost pending name parts"
        Check CurrentRevision(id) = revision, "Section switch saved a revision"
    Next
    Call UXSelectSection(8)
    Check f!cmdRoute.Top = f!cmdComments.Top And f!cmdRoute.Top < f!cmdSave.Top, "Review actions are not condensed"
    Call UXSelectSection(9)
    Check f!lstRevisions.Left = f!uxFieldSurface.Left And f!cmdCompare.Top < f!lstRevisions.Top, "History did not use its compact workspace"
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
    Check s!txtOfficialName.Top - s!labelOfficialName.Top <= 350, "Caption gap is not compact"
    Check InStr(s!labelReadinessDate.Caption, ChrW(226)) = 0, "Readiness caption contains corrupted punctuation"
    Call UXSelectSection(4)
    Check s!cboTrait33.Width = 1000 And s!cboTrait34.Top - s!cboTrait33.Top <= 700, "Trait rows still waste space"
    Check s!cboTrait33.Visible And s!cboTrait39.Visible, "All seven traits are not together on the Front Page"
    Call UXSelectSection(0)
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
    Check DCount("*", "RevisionValue", "RevisionID=" & CurrentRevision(id)) = UBound(DraftFieldKeys()) + 1, "Surface save lost snapshot fields"
    Check DLookup("MemberDoDID", "Report", "ReportID=" & id) = "1234567890", "Hidden member identifier lost"
    Check DLookup("ReportingSeniorDoDID", "Report", "ReportID=" & id) = "0987654321", "Hidden reporting senior identifier lost"
    Call UXSelectSection(5): Check s!txtOfficialComments.Visible And s!txtOfficialComments.Height >= 2000, "Narrative lacks spacious editor"
    Call UXSelectSection(4): Check s!uxTraitGuidance.Visible And s!uxTraitGuidance.Locked, "Trait guidance unavailable or editable"
    Call UXLeaveEditor("Discard")
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
    Check CLng(Forms!frmUXGroups!lstGroups.Column(4)) = 1, "Group report count is incorrect"
    Check UXOpenGroup(), "Group list did not open focused workspace"
    Check Forms!frmUXComparison!uxGroupContext = groupID, "Group workspace opened another group"
    Check InStr(Forms!frmUXComparison!lblHeading.Caption, "Surface group") > 0, "Group title missing"
    Check Forms!frmUXComparison!uxTraitGrid.Form!ReportID = id And Forms!frmUXComparison!uxTraitGrid.Form!uxOpenReport.OnClick = "=UXGridOpenReport()", "Group members lack report entry point"
    Check UXCloseComparison(), "Group workspace did not close"
    stage = "my work": DoCmd.OpenForm "frmMyWork"
    Check Forms!frmMyWork!lstWork.Visible And Not Forms!frmMyWork!lstTasks.Visible And Not Forms!frmMyWork!lstContributors.Visible, "My Work default mixes contributors and work"
    Check UXMyWorkMode("Assign"), "Assignment workspace unavailable"
    Forms!frmMyWork!uxAssignReport = id: Check UXChooseAssignmentReport(), "Named report selection failed"
    Check Forms!frmMyWork!txtExpectedRevision = CurrentRevision(id), "Assignment did not capture saved revision"
    Check UXMyWorkMode("Contributors") And Forms!frmMyWork!lstContributors.Visible And Not Forms!frmMyWork!lstWork.Visible, "Contributor workspace not focused"
    TestUXSurface = "PASS: " & count & " native surface, buffer, validation, captions and dashboard assertions": Exit Function
Failed:
    TestUXSurface = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
