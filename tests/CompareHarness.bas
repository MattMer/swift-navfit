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
    Check f!lstCompareFields.ListCount = UBound(DraftFieldKeys()) + 1, "All-fields filter wrong"
    Check CloseRevisionComparison(), "Close failed"
    Check Screen.ActiveForm.Name = "frmEditor", "Return to editor failed"
    Check Forms!frmEditor!txtTitle = "UNSAVED TITLE", "Unsaved entries lost"
    Check DLookup("Title", "Report", "ReportID=" & id) = "mixed case", "Comparison saved editor changes"
    Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Comparison changed history"
    Check OpenRevisionComparison(id, b, b), "Same revision comparison failed"
    Check Forms!frmRevisionCompare!lstCompareFields.ListCount = 0, "Same revision shows changes"
    Call CloseRevisionComparison()
    Call FixtureCloseEditor()
    other = CommitDraft(0, "FITREP", "Other", "Synthetic", Null, Null, "")
    Check WrongReport(other, a), "Cross-report comparison accepted"
    Call CommitDraft(id, "EVAL", "mixed case", "Synthetic", Null, Null, "Notes", "Empty", "")
    c = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check OpenRevisionComparison(id, a, c), "Null-empty comparison failed"
    Forms!frmRevisionCompare!lstCompareFields = "OfficialComments": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtNewValue = "[Empty text]", "Empty marker wrong"
    Call CloseRevisionComparison()
    'Simulate a migrated pre-narrative snapshot without changing the report itself:
    'an imported full copy (no ValueEncoding) without the narrative row.
    MakeFullCopy a
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

' Sparse snapshots (schema 50, L03): a save stores only fields with a value; every
' reader treats a missing field as empty. Full copies from earlier versions still read.
Public Function TestSparseRevisions() As String
    Dim id As Long, sparseRev As Long, fullRev As Long, latest As Long, key As Variant, rows As Long, stored As Long, values As Object, value As Variant
    On Error GoTo Failed
    checks = 0
    Call ConfigureWorkspace("SYNTHETIC SPARSE", "Synthetic Tester")
    id = CommitDraft(0, "EVAL", "Sparse draft", "Synthetic", Null, Null, "Notes", "Initial", "Narrative")
    sparseRev = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check Nz(DLookup("ValueEncoding", "ReportRevision", "RevisionID=" & sparseRev), 0) = 1, "New revision is not marked sparse"
    rows = DCount("*", "RevisionValue", "RevisionID=" & sparseRev)
    Check rows > 0 And rows < UBound(DraftFieldKeys()) + 1, "Sparse save wrote a row for every field"
    Check DCount("*", "RevisionValue", "RevisionID=" & sparseRev & " AND IsNullValue=True") = 0, "Sparse save stored an empty field"
    Set values = SavedFields(id)
    Check values("Title") = "Sparse draft" And IsNull(values("Trait33")) And IsNull(values("ReportingSeniorName")), "Saved fields wrong for a sparse revision"
    Check MajorRevisionHasField(sparseRev, "Trait33") And IsNull(MajorRevisionValue(sparseRev, "Trait33")), "Missing sparse field is not empty"
    Check RevisionStoredValue(sparseRev, "Title", value), "Accessor missed a stored field"
    Check value = "Sparse draft", "Accessor misread a stored field"
    Check IsNull(TraitsFromRevision(sparseRev)(0)) And IsNull(AdministrationFromRevision(sparseRev)(0)) And IsNull(ReportingSeniorFromRevision(sparseRev)(0)), "Group readers rejected a sparse revision"
    Check IsNull(PromotionStatusFromRevision(sparseRev)) And IsNull(ReportPeopleFromRevision(sparseRev)(0)) And IsNull(PartyNamesFromRevision(sparseRev)(0)), "Group readers rejected a sparse revision"
    Call AssertReportCache(id): checks = checks + 1
    ' An imported full copy: every field has a row, empty ones flagged.
    Call CommitDraft(id, "EVAL", "Full copy", "Synthetic", Null, Null, "Notes", "Second", "Narrative two")
    fullRev = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    MakeFullCopy fullRev
    Check Not RevisionIsSparse(fullRev) And DCount("*", "RevisionValue", "RevisionID=" & fullRev) = UBound(DraftFieldKeys()) + 1, "Full copy fixture incomplete"
    Call AssertReportCache(id): checks = checks + 1
    ' A full copy still reports a field it does not hold as missing.
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & fullRev & " AND FieldKey='Trait34'", dbFailOnError
    Check Not MajorRevisionHasField(fullRev, "Trait34"), "Full copy without a row read as empty"
    CurrentDb.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & fullRev & ",'Trait34',Null,True)", dbFailOnError
    ' Saving on top of a full copy writes a sparse revision; comparing them shows only real changes.
    Call CommitDraft(id, "EVAL", "Full copy", "Synthetic", Null, Null, "Notes", "Third", "Narrative changed")
    latest = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check RevisionIsSparse(latest), "Save after a full copy is not sparse"
    Check RevisionFieldStatus(fullRev, latest, "Trait33") = "Unchanged", "Empty field differs between full and sparse revisions"
    Check RevisionFieldStatus(fullRev, latest, "OfficialComments") = "Changed", "Real change hidden between full and sparse revisions"
    Check RevisionFieldStatus(sparseRev, latest, "Title") = "Changed", "Title change hidden between sparse revisions"
    ' Restoring a sparse revision brings back its values, empty fields included.
    Call RestoreDraft(id, sparseRev)
    Set values = SavedFields(id)
    Check values("Title") = "Sparse draft" And values("OfficialComments") = "Narrative" And IsNull(values("Trait33")), "Restore of a sparse revision wrong"
    Call AssertReportCache(id): checks = checks + 1
    stored = DCount("*", "RevisionValue", "RevisionID=" & DLookup("CurrentRevisionID", "Report", "ReportID=" & id))
    Check stored = rows, "Restored revision stored a different number of fields"
    TestSparseRevisions = "PASS: " & checks & " sparse revision assertions"
    Exit Function
Failed:
    TestSparseRevisions = "FAIL: " & Err.Number & " " & Err.Description
End Function

' Turn a revision into a full copy, as an earlier version saved it: every field has a
' row, empty ones flagged, and no ValueEncoding.
Private Sub MakeFullCopy(ByVal revisionID As Long)
    Dim key As Variant
    For Each key In DraftFieldKeys()
        If DCount("*", "RevisionValue", "RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(CStr(key))) = 0 Then
            CurrentDb.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & revisionID & "," & QuoteSQL(CStr(key)) & ",Null,True)", dbFailOnError
        End If
    Next
    CurrentDb.Execute "UPDATE ReportRevision SET ValueEncoding=Null,ValueCount=Null,ValueHash=Null WHERE RevisionID=" & revisionID, dbFailOnError
End Sub

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
