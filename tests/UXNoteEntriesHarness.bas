Attribute VB_Name = "UXNoteEntriesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6955, , message
    n = n + 1
End Sub
Public Function TestUXNoteEntries() As String
    Dim id As Long, rev As Long, f As Access.Form, child As Access.Form, guid As String, i As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC NOTE UI", "Note author"
    id = CommitDraft(0, "EVAL", "Note UI", "Synthetic member", #1/1/2026#, #9/30/2026#, "Legacy explanation", "Test", "")
    MigrateLegacyReportNotes
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): rev = CurrentRevision(id)
    Check UXSelectSection(7), "Notes section did not open"
    Set child = Forms!frmEditor!uxFieldSurface.Form
    Check child!uxNoteList.Visible And Not child!txtNotes.Visible, "Legacy single note editor still shown"
    Check child!uxNoteList.ListCount > 0, "Legacy entry list missing"
    Check UXOpenNote(), "Composer did not open"
    Set f = Forms!frmUXNoteComposer
    f!txtNote.SetFocus: f!txtNote.Text = "Pending note with raw keyboard input"
    Check UXNoteCheckpoint(), "Raw note checkpoint failed"
    Check CurrentRevision(id) = rev And DCount("*", "ReportNoteDraft") = 1, "Checkpoint changed revision or was lost"
    Check UXKeepNoteCheckpoint(), "Close and keep checkpoint failed"
    UXSelectSection 7
    Check child!uxNoteDrafts.ListCount = 1, "Unposted checkpoint not listed"
    child!uxNoteDrafts = child!uxNoteDrafts.ItemData(0)
    Check UXOpenNote(True), "Recovery entry failed"
    Set f = Forms!frmUXNoteComposer
    Check f!txtNote = "Pending note with raw keyboard input", "Recovered note text changed"
    Check UXPostNote(), "Posting recovered note failed"
    Check CurrentRevision(id) = rev And DCount("*", "ReportNoteDraft") = 0, "Posting changed official revision or left checkpoint"
    Check DCount("*", "ReportNote", "ReportID=" & id) = 2, "Posted note entry missing"
    For i = 0 To child!uxNoteList.ListCount - 1
        If Nz(child!uxNoteList.Column(1, i), "") = "Note author" Then child!uxNoteList = child!uxNoteList.Column(0, i)
    Next
    Check UXSelectReportNote(), "Full entry could not be selected"
    Check Nz(child!uxNoteDetails, "") = "Pending note with raw keyboard input", "Full viewer missing: count=" & child!uxNoteList.ListCount & "; selected=" & Nz(child!uxNoteList, "null") & "; first=" & Nz(child!uxNoteList.Column(1, 0), "null")
    Check UXOpenNote(), "Second composer did not open"
    Forms!frmUXNoteComposer!txtNote = "Discard me": UXNoteCheckpoint
    Check UXDiscardNote(), "Discard did not close composer"
    Check DCount("*", "ReportNoteDraft") = 0 And DCount("*", "ReportNote") = 2, "Discard altered posted notes or left pending text"
    TestUXNoteEntries = "PASS: " & n & " note entry/recovery UI assertions": Exit Function
Failed:
    TestUXNoteEntries = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function



Public Function StageNoteShutdown() As Boolean
    Dim id As Long
    ConfigureWorkspace "SYNTHETIC NOTE SHUTDOWN", "Note author"
    id = CommitDraft(0, "EVAL", "Shutdown note", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): UXSelectSection 7: UXOpenNote
    With Forms!frmUXNoteComposer!txtNote
        .SetFocus: .Text = "RAW UNSUBMITTED NOTE AT ACCESS SHUTDOWN"
    End With
    StageNoteShutdown = True
End Function
