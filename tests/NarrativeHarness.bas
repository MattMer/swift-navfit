Attribute VB_Name = "NarrativeHarness"
Option Explicit
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 500, , message
End Sub
Public Function TestNarratives() As String
    Dim kind As Variant, id As Long, first As Long, copy As Long, text As String, block As String, stage As String
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC NARRATIVE", "Synthetic Tester")
    text = "Mixed CASE; apostrophe ' and quotes ""example""" & vbCrLf & String$(6000, "x")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        stage = "Create " & kind
        id = CommitDraft(0, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes", "Narrative test", text)
        first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check StrComp(StoredNarrative(id), text, vbBinaryCompare) = 0, "Multiline narrative changed"
        Check DLookup("DraftNotes", "Report", "ReportID=" & id) = "Separate notes", "Notes mixed with narrative"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted narrative caused revision"
        stage = "Duplicate " & kind
        copy = DuplicateDraft(id)
        Check StrComp(StoredNarrative(copy), text, vbBinaryCompare) = 0, "Duplicate lost narrative"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes", "Case edit", Replace(text, "Mixed CASE", "mixed case"))
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Case-only edit not recorded"
        Call RestoreDraft(id, first)
        Check StrComp(StoredNarrative(id), text, vbBinaryCompare) = 0, "Restore lost narrative"
        stage = "Open editor " & kind
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        If kind = "EVAL" Then block = "43" Else If kind = "CHIEFEVAL" Then block = "40" Else block = "41"
        Check InStr(Forms!frmEditor!lblOfficialComments.Caption, "block " & block) > 0, "Wrong block label"
        Check StrComp(Forms!frmEditor!txtOfficialComments, text, vbBinaryCompare) = 0, "Editor failed to load narrative"
        Forms!frmEditor!txtOfficialComments = "UI narrative " & kind
        stage = "UI save " & kind
        Check SaveEditor(), "UI save failed"
        Check StoredNarrative(id) = "UI narrative " & kind, "UI save lost narrative"
        Call FixtureCloseEditor()
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes", "Clear", Null)
        Check IsNull(StoredNarrative(id)), "Null clear failed"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes", "Empty", "")
        Check Not IsNull(StoredNarrative(id)), "Empty became null"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Separate notes")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 6, "Omitted empty narrative caused revision"
    Next
    TestNarratives = "PASS: 39 narrative assertions across three forms"
    Exit Function
Failed:
    TestNarratives = "FAIL: " & stage & " " & Err.Number & " " & Err.Description
End Function


Private Function StoredNarrative(ByVal id As Long) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT OfficialComments FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    StoredNarrative = r!OfficialComments
    r.Close
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
