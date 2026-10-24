Attribute VB_Name = "UXTraitRetryHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7780, , message
    checks = checks + 1
End Sub
Public Function TestTraitRetry() As String
    Dim id As Long, other As Long, groupID As Long, route As Long, template As Long, revision As Long, db As DAO.Database
    Dim f As Access.Form, rows As Access.Form, r As DAO.Recordset, key As Variant, mark As String, failure As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC TRAIT RETRY", "Tester"
    id = CommitDraft(0, "CHIEFEVAL", "Approved", "Synthetic approved", Null, Null, "", "Test", Null, , Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"), Array("Observed", "Promotable"))
    other = DuplicateDraft(id)
    groupID = CreateSummaryGroup("Approval retry group", "CHIEFEVAL", "E7")
    AddGroupMember groupID, id, "E7": AddGroupMember groupID, other, "E7"
    template = CreateRouteTemplate("Synthetic route", CStr(MajorAuthor()), "Test")
    route = StartRoute(id, template, groupID): DecideRouteStep route, "Accepted", "Test": ApproveSavedRoute route, "Test"
    revision = CurrentRevision(id)
    Set db = CurrentDb: db.TableDefs("Report").Fields("Trait33").ValidationRule = "<> '5.0'"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
    Check UXOpenComparison(), "Open failed: " & Forms!frmMajorWorkbench!lblStatus.Caption
    Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
    Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
    rows!Trait33 = "5.0"
    Check Not UXGridEdit("Trait33", vbNo), "No reopening accepted"
    Check CurrentRevision(id) = revision And DLookup("WorkflowState", "Report", "ReportID=" & id) = "Approved", "Decline changed saved report"
    rows!Trait33 = "5.0"
    Check Not UXGridEdit("Trait33", vbYes), "Injected save failure accepted"
    Check CurrentRevision(id) = revision And DLookup("WorkflowState", "Report", "ReportID=" & id) = "Approved", "Failed save lost approval or revision"
    rows!Trait33 = "3.0"
    Check UXGridEdit("Trait33", vbYes), "Retry failed: " & f!lblStatus.Caption
    Check DLookup("Trait33", "Report", "ReportID=" & id) = "3.0", "Retried trait reverted"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft", "Approval not reopened"
    Check DLookup("RunState", "RouteRun", "RunID=" & route) = "Reopened", "Approval history lost"
    For Each key In TraitKeys()
        mark = IIf(key = "Trait33", "2.0", "5.0")
        rows.Controls(key) = mark
        Check UXGridEdit(CStr(key)), "Subsequent trait failed: " & f!lblStatus.Caption
        Check DLookup(CStr(key), "Report", "ReportID=" & id) = mark, "Subsequent trait reverted"
    Next
    Check Not IsNull(NumericTraitAverage(CurrentRevision(id))), "ITA unavailable"
    UXCloseComparison
    Check UXShowGroupMembers(groupID), "Reopen failed"
    Set rows = Forms!frmUXComparison!uxTraitGrid.Form
    Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
    Check rows!Trait33 = "2.0" And rows!Trait34 = "5.0", "Reopened comparison lost saved values"
    TestTraitRetry = "PASS: " & checks & " approved-trait rollback, retry, repeated-save and reopen assertions": Exit Function
Failed:
    failure = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
    TestTraitRetry = failure
    On Error Resume Next
    UXCloseComparison
    DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
    'This is a disposable fixture. Do not modify locked table definitions during
    'failure cleanup and obscure the original assertion/operation error.
    On Error GoTo 0
End Function
