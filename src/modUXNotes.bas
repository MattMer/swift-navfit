Attribute VB_Name = "modUXNotes"
Option Explicit
Public Function UXNotesSection(ByVal f As Access.Form, ByVal width As Long) As Boolean
 UXNotesSection = UXToolsChoose("Notes")
End Function
Public Function UXOpenNote(Optional ByVal recover As Boolean = False) As Boolean
 If Not UXToolsChoose("Notes") Then Exit Function
 If recover Then UXOpenNote = UXMessageResume() Else UXOpenNote = UXMessageBegin(False)
End Function
Public Function UXSelectReportNote() As Boolean
 UXSelectReportNote = UXMessageRead()
End Function
Public Function UXNoteDraftSelection() As Boolean
 UXNoteDraftSelection = True
End Function
Public Function UXLoadNoteComposer() As Boolean
 UXLoadNoteComposer = UXToolsChoose("Notes")
End Function
Public Function UXNoteCheckpoint() As Boolean
 UXNoteCheckpoint = UXToolsCheckpoint()
End Function
Public Function UXPostNote() As Boolean
 UXPostNote = UXMessagePost()
End Function
Public Function UXDiscardNote() As Boolean
 UXDiscardNote = UXMessageDiscard()
End Function
Public Function UXNoteUnload() As Boolean
 UXNoteUnload = UXToolsCheckpoint()
End Function
Public Function UXNoteClosed() As Boolean
 UXNoteClosed = True
End Function
Public Function UXKeepNoteCheckpoint() As Boolean
 UXKeepNoteCheckpoint = UXMessageKeep()
End Function
Public Function UXOpenNoteRecoveries() As Boolean
 If Not UXLeaveEditor() Then Exit Function
 DoCmd.OpenForm "frmUXNoteRecovery"
 Forms!frmUXNoteRecovery!lstDrafts.RowSource = "SELECT DraftGuid,UXMemberLabel(ReportID) AS MemberName,Format(CheckpointAt,'yyyy-mm-dd hh:nn') AS Checkpoint FROM ReportNoteDraft WHERE AuthorProfileID=" & MajorAuthor() & " ORDER BY CheckpointAt DESC"
 UXOpenNoteRecoveries = True
End Function
Public Function UXResumeSelectedNote() As Boolean
 Dim identity As String, id As Long
 identity = Nz(Forms!frmUXNoteRecovery!lstDrafts, "")
 id = Nz(DLookup("ReportID", "ReportNoteDraft", "DraftGuid=" & QuoteSQL(identity) & " AND AuthorProfileID=" & MajorAuthor()), 0)
 If id = 0 Then Exit Function
 If Not UXOpenReportTool(id, "Notes") Then Exit Function
 If Not UXMessageResume(identity) Then Exit Function
 DoCmd.Close acForm, "frmUXNoteRecovery", acSaveNo
 UXResumeSelectedNote = True
End Function
