Attribute VB_Name = "UXFeedbackHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6700, , message
    n = n + 1
End Sub
Public Function TestUXFeedback() As String
    Dim id As Long, rev As Long, p As Long, original As Long, snippet As Long, f As Access.Form, changes As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC FEEDBACK", "Feedback tester"
    original = MajorAuthor(): p = AddContributor("Second tester")
    DoCmd.OpenForm "frmHome"
    Forms!frmHome!uxContributor = p
    Check UXOverviewContributor(), "Overview contributor switch failed"
    Check MajorAuthor() = p, "Contributor did not change"
    Check InStr(Forms!frmHome!lblProfile.Caption, "Second tester") > 0, "Contributor header not refreshed"
    id = CommitDraft(0, "EVAL", "Feedback", "Test member", #1/1/2026#, #9/30/2026#, "", "Test", "Saved narrative")
    rev = CurrentRevision(id)
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor
    f!txtOfficialComments = "Pending narrative": UXSurfaceRefresh
    UXSelectSection 5
    Check UXOpenReportSnippets(), "Report snippet entry failed"
    Check UXReturnSnippets(), "Snippet return failed"
    Check CurrentProject.AllForms("frmEditor").IsLoaded And f!txtOfficialComments = "Pending narrative", "Snippet return lost buffer"
    Check Screen.ActiveForm.Name = "frmEditor", "Snippet close did not select report"
    Check CurrentRevision(id) = rev, "Snippet navigation saved a revision"
    f!cmdClose.SetFocus: f!txtMemberSSN = "123": UXSurfaceRefresh
    Check Not UXSaveEditor(), "Invalid SSN saved"
    Check UXLeaveEditor(), "Checkpointed invalid report could not leave"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded And CurrentRevision(id) = rev, "Invalid navigation saved or retained the editor"
    Check DLookup("MemberSSN", "RecoveryDraft", "ReportID=" & id) = "123", "Invalid navigation lost raw SSN"
    Check RestoreRecoveryDraft(DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)), "Navigation checkpoint could not recover"
    Check UXUndoPending(), "Undo failed to escape failed-save state"
    Check CurrentRevision(id) = rev, "Undo changed committed revision"
    Check Forms!frmEditor!txtOfficialComments = "Saved narrative", "Undo did not restore saved narrative"
    Check DCount("*", "RecoveryDraft", "ReportID=" & id) = 0, "Undo retained session checkpoint"
    Check UXNavigate("frmHome"), "Navigation still blocked after undo"
    Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Editor did not close after undo"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtOfficialComments = "Stale pending edits": UXSurfaceRefresh
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Other contributor saved value"
    CommitSavedFields id, CurrentRevision(id), changes, "Synthetic concurrent save"
    rev = CurrentRevision(id)
    Check Not UXSaveEditor(), "Stale save was accepted"
    Check UXUndoPending(), "Stale save could not be undone"
    Check Forms!frmEditor!txtOfficialComments = "Other contributor saved value" And CurrentRevision(id) = rev, "Undo overwrote concurrent saved work"
    Forms!frmEditor!txtOfficialComments = "Approval race pending edit": UXSurfaceRefresh
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & id
    Check Not UXSaveEditor(), "Approval race bypassed save lock"
    Check UXUndoPending(), "Approval race left editor trapped"
    Check Not Forms!frmEditor!cmdSave.Enabled, "Undo did not reload local approval lock"
    Check UXNavigate("frmHome"), "Locked report could not leave after undo"
    OpenScreen "frmSnippets"
    Check UXReturnSnippets(), "Resources snippet return failed"
    Check CurrentProject.AllForms("frmUXResources").IsLoaded, "Resources return destination missing"
    Check CurrentDb.Properties("StartupShowDBWindow") = False, "Objects pane startup preference not set"
    Check Not CurrentDb.Properties("AllowFullMenus") And Not CurrentDb.Properties("AllowSpecialKeys") And Not CurrentDb.Properties("AllowBuiltInToolbars") And Not CurrentDb.Properties("AllowBreakIntoCode"), "Lockdown startup properties not set"
    Check CurrentDb.Properties("AllowBypassKey") And CurrentDb.Properties("AppTitle") = "Swift NAVFIT", "Shift bypass or application title not set"
    Check UXHideAccessObjects(), "Native objects pane hide failed"
    TestUXFeedback = "PASS: " & n & " feedback navigation and recovery assertions": Exit Function
Failed:
    TestUXFeedback = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
