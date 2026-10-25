Attribute VB_Name = "ReportNotesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6953, , message
    n = n + 1
End Sub
Private Function RejectPost(ByVal guid As String, ByVal version As Long) As Boolean
    On Error GoTo Rejected
    PostReportNote guid, version: Exit Function
Rejected:
    RejectPost = True
End Function
Public Function TestReportNotes() As String
    Dim id As Long, rev As Long, author As Long, other As Long, guid As String, version As Long, note As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC NOTE ENTRIES", "First contributor": author = MajorAuthor(): other = AddContributor("Second contributor")
    id = CommitDraft(0, "EVAL", "Notes test", "Synthetic member", #1/1/2026#, #9/30/2026#, "Original legacy notes", "Test", "")
    rev = CurrentRevision(id)
    Check MigrateLegacyReportNotes() = 1, "Legacy note not preserved"
    Check MigrateLegacyReportNotes() = 0, "Legacy conversion duplicated notes"
    Check DLookup("ContributorName", "ReportNote", "ReportID=" & id) = "Legacy notes (author not tracked)", "Legacy author invented"
    guid = NewIdentity(): version = SaveReportNoteDraft(guid, id, 0, "Why I changed this report")
    Check version = 1 And DCount("*", "ReportNoteDraft") = 1, "Note checkpoint missing"
    Check CurrentRevision(id) = rev, "Checkpoint altered official revision"
    version = SaveReportNoteDraft(guid, id, version, "Revised local explanation")
    Check version = 2, "Checkpoint version not advanced"
    Check RejectPost(guid, 1), "Stale note posted"
    SwitchContributor other
    Check RejectPost(guid, version), "Another contributor posted draft"
    SwitchContributor author
    ' Posted notes are report messages; the compatibility result is the negated MessageID.
    note = PostReportNote(guid, version)
    Check note < 0, "Post returned a legacy NoteID"
    Check Nz(DLookup("MessageBody", "ReportMessage", "MessageID=" & -note), "") = "Revised local explanation", "Posted body changed"
    Check Nz(DLookup("ContributorName", "ReportMessage", "MessageID=" & -note), "") = "First contributor", "Attribution incorrect"
    Check DCount("*", "ReportNoteDraft") = 0, "Posted checkpoint not cleared"
    Check CurrentRevision(id) = rev, "Note posting changed report revision"
    Check RejectPost(guid, version), "Posted note duplicated"
    guid = NewIdentity(): version = SaveReportNoteDraft(guid, id, 0, " ")
    Check RejectPost(guid, version), "Blank note posted"
    DiscardReportNoteDraft guid, version
    Check DCount("*", "ReportNoteDraft") = 0, "Own draft discard failed"
    guid = NewIdentity(): version = SaveReportNoteDraft(guid, id, 0, "Recovery survives editor shutdown")
    Check DLookup("NoteText", "ReportNoteDraft", "DraftGuid=" & QuoteSQL(guid)) = "Recovery survives editor shutdown", "Recovery body missing"
    CurrentDb.Execute "UPDATE Report SET IsArchived=True WHERE ReportID=" & id, dbFailOnError
    Check RejectPost(guid, version), "Archived report accepted new note"
    Check DCount("*", "ReportNote", "ReportID=" & id) = 1 And DCount("*", "ReportMessage", "ReportID=" & id) = 1, "Rejected actions changed committed entries"
    Check DLookup("DraftNotes", "Report", "ReportID=" & id) = "Original legacy notes", "Legacy saved field changed"
    TestReportNotes = "PASS: " & n & " attributed notes/checkpoint assertions": Exit Function
Failed:
    TestReportNotes = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
