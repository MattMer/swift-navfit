Attribute VB_Name = "UXResponsiveIdentityHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise vbObjectError + 5400, , message
    checks = checks + 1
End Sub
Private Sub Fits(ByVal f As Access.Form, ByVal width As Long)
    Dim c As Access.Control
    For Each c In f.Controls
        If c.Visible And c.ControlType <> acRectangle Then
            Check c.Left >= 0 And c.Left + c.Width <= width, f.Name & "/" & c.Name & " exceeds viewport " & width & " (left " & c.Left & ", width " & c.Width & ")"
        End If
    Next
End Sub
Public Function TestUXResponsiveIdentity() As String
    Dim id As Long, rev As Long, groupID As Long, meeting As Long, reviewer As Long
    Dim f As Access.Form, child As Access.Form, key As Variant, name As Variant, changes As Object, stage As String, x As Long, y As Long, w As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC RESPONSIVE TEST", "Layout tester"
    stage = "identity setup"
    id = CommitDraft(0, "EVAL", "Identity report", "Legacy display", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialName", "UPDATED OFFICIAL NAME"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Official name test")
    rev = CurrentRevision(id)
    Check UXMemberLabel(id) = "UPDATED OFFICIAL NAME", "Official name is not the primary identity"
    Check UXRevisionMemberLabel(rev) = "UPDATED OFFICIAL NAME", "Snapshot identity ignores official name"
    Check UXIdentity(Null, "Local fallback") = "Local fallback", "Local fallback missing"
    Check UXIdentity("   ", "  ") = "Unnamed member", "Empty name fallback missing"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor: Set child = f!uxFieldSurface.Form
    Check InStr(f!uxReport.Caption, "UPDATED OFFICIAL NAME") = 1, "Report header still prefers local name"
    Check f!uxAlias.Caption = "Local name: Legacy display", "Independent local name not identified"
    f!uxFieldSurface.SetFocus: child!txtOfficialName.SetFocus
    child!txtOfficialName.Text = "PENDING OFFICIAL NAME": child!txtOfficialName.SelStart = 4: child!txtOfficialName.SelLength = 3
    Check UXResizeEditor(10000, 6800), "Narrow resize rejected"
    Check Screen.ActiveControl.Name = "txtOfficialName", "Resize changed active field: " & Screen.ActiveControl.Name
    Check child!txtOfficialName.SelStart = 4 And child!txtOfficialName.SelLength = 3, "Resize lost caret/selection"
    Check child!txtOfficialName.Text = "PENDING OFFICIAL NAME", "Resize lost active text"
    Check CurrentRevision(id) = rev, "Resize saved a revision"
    stage = "fixed page viewport"
    x = child!txtOfficialName.Left: y = child!txtOfficialName.Top: w = child!txtOfficialName.Width
    For Each key In Array(7500, 10000, 25000)
        Check UXResizeEditor(CLng(key), 6800), "Viewport resize rejected: " & key
        Check f!uxRail.Visible And f!uxNav0.Visible, "Navigation rail hidden at " & key
        Check f!uxFieldSurface.Left > f!uxRail.Left + f!uxRail.Width, "Report overlaps navigation rail at " & key
        Check child!txtOfficialName.Left = x And child!txtOfficialName.Top = y And child!txtOfficialName.Width = w, "Resize reflowed fields at " & key
        Check child!txtOfficialName.Text = "PENDING OFFICIAL NAME" And CurrentRevision(id) = rev, "Resize lost pending text or saved at " & key
    Next
    Check UXSaveEditor(), "Pending name could not save"
    Check DLookup("MemberName", "Report", "ReportID=" & id) = "Legacy display", "Official edit overwrote independent local name"
    Check UXMemberLabel(id) = "PENDING OFFICIAL NAME", "Saved identity not refreshed"
    Call UXLeaveEditor()
    stage = "list identities"
    DoCmd.OpenForm "frmLibrary": RefreshLibrary
    Check Forms!frmLibrary!lstReports.Column(3, IIf(Forms!frmLibrary!lstReports.ColumnHeads, 1, 0)) = "PENDING OFFICIAL NAME", "Reports list identity mismatch"
    DoCmd.OpenForm "frmHome": UXDashboardRefresh
    Check Forms!frmHome!uxRecentReports.Column(1, 1) = "PENDING OFFICIAL NAME", "Overview identity mismatch"
    groupID = CreateSummaryGroup("Layout group", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID: Forms!frmMajorWorkbench!uxGroupContext = groupID
    Check UXGroupNavigate("frmMajorWorkbench", 0), "Group members failed"
    Check Forms!frmUXComparison!uxTraitGrid.Form!Member = "PENDING OFFICIAL NAME", "Group identity mismatch"
    meeting = CreateComparisonSet(FreezeSummaryGroup(groupID, "Synthetic comparison"), "Identity comparison")
    DoCmd.Close acForm, "frmUXComparison", acSaveNo
    DoCmd.OpenForm "frmUXComparison", , , , , , CStr(meeting)
    Check Forms!frmUXComparison!uxTraitGrid.Form!Member = "PENDING OFFICIAL NAME", "Comparison identity mismatch"
    DoCmd.OpenForm "frmMyWork": UXMyWorkMode "Assign"
    Check InStr(Forms!frmMyWork!uxAssignReport.RowSource, "UXIdentity") > 0, "Assignment selector identity mismatch"
    stage = "shared workspace bounds"
    For Each name In Array("frmHome", "frmLibrary", "frmMajorWorkbench", "frmMyWork", "frmUXComparison", "frmUXSeniors", "frmUXSeniorEditor", "frmUXGroupRSCA", "frmUXGroups", "frmUXBulk")
        If Not CurrentProject.AllForms(CStr(name)).IsLoaded Then
            If name = "frmUXSeniorEditor" Then
                DoCmd.OpenForm CStr(name), , , , , , "0"
            Else
                DoCmd.OpenForm CStr(name), , , , , , CStr(groupID)
            End If
        End If
        Set f = Forms(CStr(name))
        For Each key In Array(10000, 7500, 13800)
            Check UXResizeWorkspace(CStr(name), CLng(key), 6800), "Workspace resize failed: " & name & " " & UXViewportFailure()
            Fits f, CLng(key)
            Check f.ScrollBars = 2, "Safe vertical access missing: " & name
        Next
    Next
    stage = "stacked page slots"
    'Hidden pages and views stacked in a slot must not shrink one row's fields.
    Check UXResizeWorkspace("frmUXSeniorEditor", 18000, 6800), "Wide senior editor resize failed: " & UXViewportFailure()
    Set f = Forms!frmUXSeniorEditor
    For Each key In Array("rsReportingSeniorDesignator", "rsReportingSeniorUIC")
        Check f.Controls(CStr(key)).Width = f!rsReportingSeniorName.Width And f.Controls(CStr(key)).Left = f!rsReportingSeniorName.Left, "Left column width differs: " & key
    Next
    For Each key In Array("rsReportingSeniorTitle", "rsReportingSeniorSSN")
        Check f.Controls(CStr(key)).Width = f!rsReportingSeniorGrade.Width And f.Controls(CStr(key)).Left = f!rsReportingSeniorGrade.Left, "Right column width differs: " & key
    Next
    stage = "tall member grid"
    'The members grid shows twelve one-line rows by design and takes spare window height in
    'whole rows, with the controls under it moving down.
    Set f = Forms!frmUXComparison
    Check UXResizeWorkspace("frmUXComparison", 18000, 6800), "Short comparison resize failed: " & UXViewportFailure()
    Check f!uxTraitGrid.Height = 6150, "Members grid is not twelve rows: " & f!uxTraitGrid.Height
    Check f!uxGridAverage.Top >= f!uxTraitGrid.Top + f!uxTraitGrid.Height, "Grid average overlaps the members grid"
    Check UXResizeWorkspace("frmUXComparison", 18000, 16000), "Tall comparison resize failed: " & UXViewportFailure()
    Check f!uxTraitGrid.Height > 6150 And (f!uxTraitGrid.Height - 6150) Mod 430 = 0, "Members grid did not grow by whole rows: " & f!uxTraitGrid.Height
    Check f!uxGridAverage.Top >= f!uxTraitGrid.Top + f!uxTraitGrid.Height, "Grid average did not move under the grown grid"
    Check f!uxCurrentExport.Top + f!uxCurrentExport.Height <= 16000, "Group footer pushed past a tall window"
    Check UXResizeWorkspace("frmUXComparison", 18000, 30000), "Very tall comparison resize failed: " & UXViewportFailure()
    Check f!uxTraitGrid.Height = 6150 + 6450, "Members grid growth is not capped: " & f!uxTraitGrid.Height
    UXRestoreViewport f
    stage = "unchanged viewport cache"
    UXResizeWorkspace "frmLibrary", 13800, 6800
    key = UXViewportPasses()
    Check UXResizeWorkspace("frmLibrary", 13800, 6800), "Cached viewport must succeed"
    Check UXViewportPasses() = CLng(key), "Same dimensions repeated geometry writes"
    UXResizeWorkspace "frmLibrary", 13800, 7000
    Check UXViewportPasses() = CLng(key) + 1, "Changed height did not lay out"
    UXRestoreViewport Forms!frmLibrary
    UXResizeWorkspace "frmLibrary", 13800, 7000
    Check UXViewportPasses() = CLng(key) + 2, "Restored mode geometry was not invalidated"
    DoCmd.SelectObject acForm, "frmHome", False
    UXResizeWorkspace "frmLibrary", 7500, 6800
    key = Forms!frmLibrary!uxContext.Width
    UXLoadContext
    Check Forms!frmLibrary!uxContext.Width = CLng(key), "Context refresh resized hidden workspace"
    TestUXResponsiveIdentity = "PASS: " & checks & " fixed-page viewport, workspace layout, focus and independent identity assertions": Exit Function
Failed:
    TestUXResponsiveIdentity = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
