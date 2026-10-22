Attribute VB_Name = "PersistenceSeedHarness"
Option Explicit
'Uses only APIs present in the real schema-27 development build as well as 28.
'Trusted synthetic setup is separate from the read-only production importer.
Public Function SeedPersistenceTransfer() As String
    Dim id As Long, sid As Long, baseline As Long, groupID As Long, f As Object, senior As Object, key As Variant, rev As Long, domainID As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC TRANSFER", "Tester"
    id = CommitDraft(0, "EVAL", "Preserved report", "Synthetic", #1/1/2026#, #9/30/2026#, "Saved notes")
    Set senior = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys(): senior.Add CStr(key), Null: Next
    senior("ReportingSeniorSSN") = "000000001": senior("ReportingSeniorName") = "SYNTHETIC"
    sid = RSSaveSenior(0, "Synthetic senior", senior)
    Set f = CreateObject("Scripting.Dictionary")
    For Each key In senior.Keys: f.Add CStr(key), senior(key): Next
    f.Add "ObservationStatus", "Observed"
    For Each key In TraitKeys(): f.Add CStr(key), "4.0": Next
    f.Add "CounselingDate", "NOT PERF": f.Add "CounselorName", "Explanation retained"
    f.Add "RetentionYes", "Yes": f.Add "RetentionNo", "Yes"
    CommitSavedFields id, CurrentRevision(id), f, "Synthetic source snapshot"
    groupID = CreateSummaryGroup("Synthetic completed group", "EVAL", "E6"): AddGroupMember groupID, id, "E6"
    baseline = RSAddBaseline(sid, "E6", #12/31/2025#, 4, 7, 28, "Synthetic observed baseline")
    RSSetGroupBaseline groupID, sid, baseline, True
    SetGroupWorkflow groupID, GroupStamp(groupID), "Completed"
    SetGroupWorkflow groupID, GroupStamp(groupID), "In-Work"
    SetFieldDefault "Report", id, "CounselingDate", "NOT REQ"
    rev = CurrentRevision(id)
    CurrentDb.Execute "UPDATE Report SET DateReported='LEGACY' WHERE ReportID=" & id, dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='LEGACY',IsNullValue=False WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
    If DLookup("SchemaVersion", "AppMetadata") >= 28 Then
        CurrentDb.Execute "DELETE FROM RevisionDate WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
        Application.Run "WriteRevisionFacts", rev
    End If
    domainID = CommitDraft(0, "EVAL", "Legacy domain fixture", "Synthetic", Null, Null, "")
    rev = CurrentRevision(domainID)
    If DLookup("SchemaVersion", "AppMetadata") >= 28 Then
        CurrentDb.Execute "UPDATE Report SET LegacyFormKind='LEGACY',LegacyWorkflowNull=True WHERE ReportID=" & domainID, dbFailOnError
    End If
    CurrentDb.Execute "UPDATE Report SET FormKind='LEGACY',WorkflowState=Null WHERE ReportID=" & domainID, dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='LEGACY',IsNullValue=False WHERE RevisionID=" & rev & " AND FieldKey='FormKind'", dbFailOnError
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtStart = "raw pending date": Forms!frmEditor!txtNotes = "Uncommitted recovery retained"
    If Not CheckpointEditor() Then Err.Raise 5, , "Recovery fixture failed"
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
    SeedPersistenceTransfer = "PASS: synthetic revisions, posting, reopened state, conflicting flags, legacy dates/domains, defaults and raw recovery source": Exit Function
Failed:
    SeedPersistenceTransfer = "FAIL: " & Err.Number & " " & Err.Description
End Function
