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
            Check c.Left >= 0 And c.Left + c.Width <= width, f.Name & "/" & c.Name & " exceeds viewport"
        End If
    Next
End Sub
Public Function TestUXResponsiveIdentity() As String
    Dim id As Long, rev As Long, groupID As Long, meeting As Long, reviewer As Long
    Dim f As Access.Form, child As Access.Form, c As Access.Control, key As Variant, name As Variant, section As Long, changes As Object, row As Long, stage As String
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
    Check f!uxCompactNav.Visible And f!uxCompactSection.Visible And Not f!uxNav0.Visible, "Narrow navigation not collapsed"
    Check f!uxFieldSurface.Left = 250 And f!uxFieldSurface.Width = 9400, "Narrow panel retains wide geometry"
    Check f!cmdSave.Top + f!cmdSave.Height <= 6800, "Save is outside narrow viewport"
    stage = "section layouts"
    For Each key In Array(10000, 9000, 7500)
        For section = 0 To 9
            Check UXSelectSection(section), "Section cannot open"
            Check UXResizeEditor(CLng(key), 6800), "Compact section resize failed"
            Fits f, CLng(key)
            For Each c In child.Controls
                If c.Visible Then Check c.Left + c.Width <= child.Width, "Child field extends horizontally"
            Next
        Next
    Next
    Check UXSelectSection(0), "Return to member section failed"
    Check UXResizeEditor(7500, 4500), "Short viewport resize rejected"
    Check f.ScrollBars = 2 And f.Section(acDetail).Height > 4500, "Short viewport lacks vertical fallback"
    Check UXResizeEditor(25000, 13500), "Wide layout restore failed"
    Check f!uxNav0.Visible And Not f!uxCompactSection.Visible And f!uxFieldSurface.Left = 5900, "Wide navigation not restored"
    Check f!cmdSave.Top = 12200, "Wide footer not restored"
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
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
    UXWorkbenchMode "Members"
    Check Forms!frmMajorWorkbench!lstReports.Column(1, IIf(Forms!frmMajorWorkbench!lstReports.ColumnHeads, 1, 0)) = "PENDING OFFICIAL NAME", "Group identity mismatch"
    meeting = CreateComparisonSet(FreezeSummaryGroup(groupID, "Synthetic comparison"), "Identity comparison")
    DoCmd.OpenForm "frmUXComparison", , , , , , CStr(meeting)
    Check Forms!frmUXComparison!lstTraits.Column(1, IIf(Forms!frmUXComparison!lstTraits.ColumnHeads, 1, 0)) = "PENDING OFFICIAL NAME", "Comparison identity mismatch"
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
    TestUXResponsiveIdentity = "PASS: " & checks & " responsive layout, focus and independent identity assertions": Exit Function
Failed:
    TestUXResponsiveIdentity = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
