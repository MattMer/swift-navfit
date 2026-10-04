Attribute VB_Name = "CompareHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 850, , message
    checks = checks + 1
End Sub
Private Function WrongReport(ByVal reportID As Long, ByVal revisionID As Long) As Boolean
    On Error GoTo Rejected
    Call OpenRevisionComparison(reportID, revisionID)
    Exit Function
Rejected:
    WrongReport = InStr(Err.Description, "belong") > 0
End Function
Public Function TestComparison() As String
    Dim id As Long, other As Long, a As Long, b As Long, c As Long, text As String
    Dim f As Form, initial As Long
    On Error GoTo Failed
    checks = 0
    Call ConfigureWorkspace("SYNTHETIC COMPARISON", "Synthetic Tester")
    id = CommitDraft(0, "EVAL", "Mixed Case", "Synthetic", Null, Null, "Notes", "Initial", Null)
    a = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    text = "Narrative " & vbCrLf & String$(7000, "x")
    Call CommitDraft(id, "EVAL", "mixed case", "Synthetic", Null, Null, "Notes", "Second", text)
    b = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check RevisionFieldStatus(a, b, "Title") = "Changed", "Case edit invisible"
    Check RevisionFieldStatus(a, b, "OfficialComments") = "Changed", "Narrative edit invisible"
    Check RevisionFieldStatus(a, b, "MemberName") = "Unchanged", "Unchanged value misclassified"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtTitle = "UNSAVED TITLE"
    Forms!frmEditor!lstRevisions = a
    Check CompareSelectedRevision(), "Editor compare action failed"
    Check Screen.ActiveForm.Name = "frmRevisionCompare", "Comparison is not active"
    Set f = Forms!frmRevisionCompare
    Check f!lstCompareFields.ListCount = 2, "Changed filter wrong"
    f!lstCompareFields = "OfficialComments": Call LoadComparisonValues()
    Check f!txtOldValue = "[NULL]", "Null marker wrong"
    Check StrComp(f!txtNewValue, text, vbBinaryCompare) = 0, "Long narrative display changed"
    Check f!txtOldValue.Locked And f!txtNewValue.Locked, "Comparison values editable"
    f!chkChangedOnly = False: Call RefreshComparisonFields()
    Check f!lstCompareFields.ListCount = 78, "All-fields filter wrong"
    Check CloseRevisionComparison(), "Close failed"
    Check Screen.ActiveForm.Name = "frmEditor", "Return to editor failed"
    Check Forms!frmEditor!txtTitle = "UNSAVED TITLE", "Unsaved entries lost"
    Check DLookup("Title", "Report", "ReportID=" & id) = "mixed case", "Comparison saved editor changes"
    Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Comparison changed history"
    Check OpenRevisionComparison(id, b, b), "Same revision comparison failed"
    Check Forms!frmRevisionCompare!lstCompareFields.ListCount = 0, "Same revision shows changes"
    Call CloseRevisionComparison()
    DoCmd.Close acForm, "frmEditor", acSaveNo
    other = CommitDraft(0, "FITREP", "Other", "Synthetic", Null, Null, "")
    Check WrongReport(other, a), "Cross-report comparison accepted"
    Call CommitDraft(id, "EVAL", "mixed case", "Synthetic", Null, Null, "Notes", "Empty", "")
    c = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check OpenRevisionComparison(id, a, c), "Null-empty comparison failed"
    Forms!frmRevisionCompare!lstCompareFields = "OfficialComments": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtNewValue = "[Empty text]", "Empty marker wrong"
    Call CloseRevisionComparison()
    'Simulate a migrated pre-narrative snapshot without changing the report itself.
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & a & " AND FieldKey='OfficialComments'", dbFailOnError
    Check OpenRevisionComparison(id, a, c), "Legacy comparison failed"
    Forms!frmRevisionCompare!lstCompareFields = "OfficialComments": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy absence hidden"
    Call CloseRevisionComparison()
    TestComparison = "PASS: " & checks & " revision comparison assertions"
    Exit Function
Failed:
    TestComparison = "FAIL: " & Err.Number & " " & Err.Description
End Function
