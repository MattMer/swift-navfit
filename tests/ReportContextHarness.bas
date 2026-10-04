Attribute VB_Name = "ReportContextHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2550, , message
    checks = checks + 1
End Sub
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function Seed() As Variant
    Seed = Array("Yes", "No", "Yes", "Yes", "Yes", "No", "No", "PRIMARY DUTY", "Duties" & vbCrLf & String$(3000, "D"))
End Function
Private Function SaveContext(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveContext = CommitDraft(id, kind, "Context", "Synthetic", Null, Null, "Local notes", "Context test", "Official comments", , , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveContext(id, "FITREP", values): Exit Function
Failed:
    Rejected = True
End Function
Private Function RestoreRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    RestoreRejected = True
End Function
Private Function ReviewRejected(ByVal task As Long) As Boolean
    On Error GoTo Failed
    Call FinishReview(task, "Accepted", ""): Exit Function
Failed:
    ReviewRejected = True
End Function
Public Function TestReportContext() As String
    Dim kind As Variant, panel As Variant, keys As Variant, key As Variant, values As Variant, bad As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, i As Variant, before As Long, r As DAO.Recordset
    Dim screen As String, chosen As String, panelKeys As Variant, initialValues As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Context tests", "Synthetic reviewer"): keys = ReportContextKeys()
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SaveContext(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = 49, "Incomplete snapshot"
        For i = 0 To 8
            Check Stored(id, CStr(keys(i))) = values(i), "Saved context changed"
        Next
        Check Stored(id, "OfficialComments") = "Official comments" And Stored(id, "DraftNotes") = "Local notes", "Duties replaced narrative/notes"
        Call CommitDraft(id, CStr(kind), "Context", "Synthetic", Null, Null, "Local notes", "Omitted", "Official comments")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted context changed history"
        copy = DuplicateDraft(id)
        Check Stored(copy, "OccasionSpecial") = "Yes" And Stored(copy, "ReportRegular") = "Yes" And Stored(copy, "DutyDescription") = values(8), "Duplicate lost features"
        For Each i In Array(0, 4, 7)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            If i = 7 Then values(i) = "OTHER DUTY" Else values(i) = "No"
            Call SaveContext(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewRejected(task), "Feature edit left review current"
            Call FinishReview(task, "Cancelled", "Updated context")
            Check RevisionFieldStatus(first, second, CStr(keys(i))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(i): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(i), "Compared feature value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first)
        Check Stored(id, "OccasionPeriodic") = "Yes" And Stored(id, "ReportRegular") = "Yes" And Stored(id, "PrimaryDuty") = "PRIMARY DUTY", "Restore lost features"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        For Each panel In Array("Occasions", "Categories", "Duties")
            screen = "frmContext" & panel
            panelKeys = ReportContextPanelKeys(CStr(panel)): key = panelKeys(0)
            chosen = CStr(key)
            Check OpenReportContext(CStr(panel)), "Panel open failed"
            Check Forms(screen).Controls("txt" & chosen) = Stored(id, chosen), "Panel load wrong"
            Forms(screen).Controls("txt" & chosen) = "CANCEL ME"
            Check CancelReportContext(CStr(panel)), "Panel cancel failed"
            Check Forms!frmEditor.Controls("txt" & chosen) = Stored(id, chosen), "Cancel changed pending values"
            Call OpenReportContext(CStr(panel))
            If panel = "Duties" Then Forms(screen).Controls("txt" & chosen) = "OTHER DUTY" Else Forms(screen).Controls("txt" & chosen) = "No"
            Check ApplyReportContext(CStr(panel)), "Panel apply failed"
            Check Forms!frmEditor.Controls("txt" & chosen) <> Stored(id, chosen), "Apply committed or lost edit"
            Check DLookup(chosen, "RecoveryDraft", "ReportID=" & id) = Forms!frmEditor.Controls("txt" & chosen), "Apply failed to checkpoint"
        Next
        initialValues = Seed()
        Check Forms!frmEditor!txtDutyDescription = initialValues(8), "Applying another panel lost memo"
        Check SaveEditor(), "Context editor save failed"
        Check Stored(id, "OccasionPeriodic") = "No" And Stored(id, "ReportRegular") = "No" And Stored(id, "PrimaryDuty") = "OTHER DUTY", "Save lost panels"
        DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    values = Seed(): values(7) = String$(14, "X"): values(8) = String$(12000, "L") & vbCrLf & "TAIL"
    Call SaveContext(id, "FITREP", values)
    Check Stored(id, "PrimaryDuty") = values(7) And Stored(id, "DutyDescription") = values(8), "Boundary/memo truncated"
    values(7) = values(7) & "X": Check Rejected(id, values), "Overlength primary duty accepted"
    before = DCount("*", "ReportRevision")
    For i = 0 To 6
        values = Seed(): values(i) = "yes": Check Rejected(id, values), "Case variant flag accepted"
        values(i) = True: Check Rejected(id, values), "Boolean code accepted"
    Next
    For Each bad In Array(123, "DUTY" & vbCrLf, "DUTY" & vbTab)
        values = Seed(): values(7) = bad: Check Rejected(id, values), "Invalid primary duty accepted"
    Next
    values = Seed(): values(8) = 123: Check Rejected(id, values), "Nontext memo accepted"
    Check Rejected(id, Array(Null)), "Malformed context accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Array(Null, "", "No", Null, "", "No", Null, "", Null): Call SaveContext(id, "FITREP", values)
    Check IsNull(Stored(id, "OccasionPeriodic")) And Not IsNull(Stored(id, "OccasionDetachment")) And Stored(id, "OccasionFrocking") = "No", "Unentered/empty/No collapsed"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'DutyDescription'"
    Check Rejected(id, Seed()), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id, "OccasionPeriodic")) And DCount("*", "ReportRevision") = before, "Late failure did not roll back all features"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Call OpenReportContext("Occasions"): Forms!frmContextOccasions!txtOccasionPeriodic = "PENDING INVALID"
    Call ApplyReportContext("Occasions")
    Call OpenReportContext("Categories"): Forms!frmContextCategories!txtReportConcurrent = "Yes"
    Call ApplyReportContext("Categories")
    Call OpenReportContext("Duties"): Forms!frmContextDuties!txtPrimaryDuty = "PENDING OVERLONG DUTY"
    Forms!frmContextDuties!txtDutyDescription = String$(4000, "R") & vbCrLf & "RECOVER"
    Call ApplyReportContext("Duties"): DoCmd.Close acForm, "frmEditor", acSaveNo
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtOccasionPeriodic = "PENDING INVALID" And Forms!frmEditor!txtReportConcurrent = "Yes", "Recovered flags wrong"
    Check Forms!frmEditor!txtPrimaryDuty = "PENDING OVERLONG DUTY" And Len(Forms!frmEditor!txtDutyDescription) > 4000, "Raw duty recovery truncated"
    Check IsNull(Stored(id, "OccasionPeriodic")) And IsNull(Stored(id, "DutyDescription")), "Recovery auto committed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    For Each panel In Array("Occasions", "Categories", "Duties")
        Call OpenReportContext(CStr(panel)): Check Not Forms("frmContext" & panel)!cmdApply.Enabled, "Archived panel can apply"
        Call CancelReportContext(CStr(panel))
    Next
    DoCmd.Close acForm, "frmEditor", acSaveNo: Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys
        CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError
    Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=13 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "DutyDescription": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented context": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='DutyDescription'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "DutyDescription": r!IsNullValue = True: r.Update: r.Close
    TestReportContext = "PASS: " & checks & " report context assertions": Exit Function
Failed:
    TestReportContext = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyReportContextRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyReportContextRecovery = (Forms!frmEditor!txtOccasionPeriodic = "PENDING INVALID" And Forms!frmEditor!txtReportConcurrent = "Yes" And Forms!frmEditor!txtPrimaryDuty = "PENDING OVERLONG DUTY" And Len(Forms!frmEditor!txtDutyDescription) > 4000)
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Function
