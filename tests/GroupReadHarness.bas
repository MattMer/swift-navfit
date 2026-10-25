Attribute VB_Name = "GroupReadHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8760, , message
    checks = checks + 1
End Sub
Private Function AverageError(ByVal g As Long) As Long
    Dim value As Variant
    On Error GoTo Failed
    value = UXGroupAverage(g)
    Exit Function
Failed:
    AverageError = Err.Number
End Function
Private Function CommittedReadError(ByVal g As Long) As Long
    Dim result As Object, rows As DAO.Recordset
    On Error GoTo Failed
    Set result = RSGroupPSG(g)
    Set rows = RSGroupRows(g): rows.Close
    Exit Function
Failed:
    CommittedReadError = Err.Number
End Function
Private Sub Compare(ByVal g As Long)
    Dim currentAverage As Variant, fresh As Object, sharedResult As Object, key As Variant, rows As DAO.Recordset, revisions As Long, snapshots As Long, text As String
    revisions = DCount("*", "ReportRevision"): snapshots = DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet")
    currentAverage = UXGroupAverage(g)
    Set fresh = RSGroupPSG(g)
    Set rows = RSGroupRows(g)
    If Not rows.EOF Then rows.MoveLast: rows.MoveNext
    Set sharedResult = RSGroupPSG(g, rows)
    For Each key In fresh.Keys
        Check SameFieldValue(fresh(key), sharedResult(key)), "Shared PSG result differs: " & key
    Next
    text = UXGroupCalculationText(g, rows)
    Check Len(text) > 0, "Current calculation context is empty"
    Check rows.EOF, "Shared input cursor moved"
    rows.Close
    Check DCount("*", "ReportRevision") = revisions, "Calculation wrote revisions"
    Check DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet") = snapshots, "Calculation wrote snapshots"
End Sub
Public Function TestGroupReads() As String
    Dim g As Long, id As Long, other As Long, senior As Long, baseline As Long, values As Object, changes As Object, key As Variant, rev As Long, result As Object, f As Access.Form
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC GROUP READS", "Tester"
    g = CreateSummaryGroup("Read group", "EVAL", "E5")
    Compare g
    Set values = RSSeniorValues(0): values("ReportingSeniorSSN") = "000123456": values("ReportingSeniorName") = "SYNTHETIC RS"
    senior = RSSaveSenior(0, "Synthetic senior", values)
    baseline = RSAddBaseline(senior, "E5", #1/1/2026#, 4, 100, 400, "Exact fixture")
    RSSetGroupBaseline g, senior, baseline, False: Compare g
    RSSetGroupBaseline g, senior, baseline, True: Compare g
    id = CommitDraft(0, "EVAL", "First", "Member one", Null, Null, "")
    other = CommitDraft(0, "EVAL", "Second", "Member two", Null, Null, "")
    Set changes = CreateObject("Scripting.Dictionary")
    For Each key In TraitKeys(): changes(CStr(key)) = "3.0": Next
    changes("ObservationStatus") = "Observed": changes("ReportingSeniorSSN") = "000123456"
    CommitSavedFields id, CurrentRevision(id), changes, "Fixture"
    CommitSavedFields other, CurrentRevision(other), changes, "Fixture"
    AddGroupMember g, other, "E5": AddGroupMember g, id, "E5"
    Compare g
    baseline = RSAddBaseline(senior, "E5", #2/1/2026#, 4.1, 100, Null)
    RSSetGroupBaseline g, senior, baseline, True: Compare g
    baseline = RSAddBaseline(senior, "E5", #3/1/2026#, Null, 0, 0)
    RSSetGroupBaseline g, senior, baseline, True: Compare g
    changes.RemoveAll: changes("Trait33") = Null: changes("Trait34") = "NOB"
    CommitSavedFields id, CurrentRevision(id), changes, "Partial traits": Compare g
    Set result = RSGroupPSG(g): Check result("Partial"), "Partial flag missing"
    changes.RemoveAll: changes("ObservationStatus") = "NOB"
    CommitSavedFields id, CurrentRevision(id), changes, "Whole NOB": Compare g
    CommitSavedFields other, CurrentRevision(other), changes, "All NOB": Compare g
    ArchiveDraft id, True: Compare g
    ArchiveDraft id, False
    changes.RemoveAll: changes("ObservationStatus") = "Observed": changes("ReportingSeniorSSN") = "111111111"
    CommitSavedFields id, CurrentRevision(id), changes, "Senior mismatch": Compare g
    changes("ReportingSeniorSSN") = "000123456"
    CommitSavedFields id, CurrentRevision(id), changes, "Restore senior": Compare g
    ' Deliberate scratch divergence: the average reads revision traits; PSG and group rows read
    ' the derived Report cache, which now rejects a value that differs from its revision.
    CurrentDb.Execute "UPDATE Report SET Trait35='5.0' WHERE ReportID=" & id, dbFailOnError
    Check AverageError(g) = 0, "Revision-based average read the derived cache"
    Check CommittedReadError(g) = vbObjectError + 8903, "Derived cache divergence not rejected"
    CurrentDb.Execute "UPDATE Report SET Trait35='3.0' WHERE ReportID=" & id, dbFailOnError
    Compare g
    rev = CurrentRevision(id)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & rev & " AND FieldKey='Trait39'", dbFailOnError
    Check AverageError(g) = vbObjectError + 4007, "Missing-field fixture not rejected"
    CurrentDb.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & rev & ",'Trait39','3.0',False)", dbFailOnError
    Compare g
    DoCmd.OpenForm "frmMajorWorkbench": Set f = Forms!frmMajorWorkbench
    f!cboGroup = g
    Check UXFocusedGroupMode("Review"), "Workbench refresh failed"
    Check f!uxSeniorCaption.Caption = UXGroupSeniorCaption(g), "Reused footer PSG differs"
    Check RSOpenGroupCalculator(g), "Calculator open failed"
    Check RSShowGroupCalculation(), "Shared calculator refresh failed"
    Check Forms!frmUXGroupRSCA!txtCalculationContext = UXGroupCalculationText(g), "Displayed context differs"
    TestGroupReads = "PASS: " & checks & " current group calculation, sharedResult cursor, history and UI checks": Exit Function
Failed:
    TestGroupReads = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
