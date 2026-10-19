Attribute VB_Name = "modUXNotes"
Option Explicit
Private noteGuid As String
Private noteVersion As Long
Private noteReport As Long
Private noteAuthor As Long
Private noteFinished As Boolean
Public Function UXNotesSection(ByVal f As Access.Form, ByVal width As Long) As Boolean
    Dim id As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    With f!uxNoteList
        .RowSourceType = "Table/Query": .ColumnCount = 3
        .Visible = True: .Left = 200: .Top = 200: .Width = width: .Height = 1750
        .RowSource = "SELECT NoteID,ContributorName AS [Contributor],Left(NoteText,240) AS [Note] FROM ReportNote WHERE ReportID=" & id & " ORDER BY NoteID DESC"
        .ColumnWidths = "0;1800;" & CStr(width - 1900): .Requery
    End With
    With f!uxNoteDetails: .Visible = True: .Left = 200: .Top = 2100: .Width = width: .Height = 1650: .Value = Null: End With
    With f!uxAddNote: .Visible = True: .Left = 200: .Top = 3900: .Width = width: .Height = 450: .Enabled = id > 0 And Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True): End With
    With f!uxNoteDrafts
        .ColumnCount = 2
        .Visible = True: .Left = 200: .Top = 4500: .Width = width: .Height = 1000
        .RowSource = "SELECT DraftGuid,'Unposted note - ' & Format(CheckpointAt,'yyyy-mm-dd hh:nn') AS Checkpoint FROM ReportNoteDraft WHERE ReportID=" & id & " AND AuthorProfileID=" & MajorAuthor() & " ORDER BY CheckpointAt DESC"
        .ColumnWidths = "0;" & CStr(width - 100): .Requery
    End With
    With f!uxRecoverNote: .Visible = True: .Left = 200: .Top = 5650: .Width = width: .Height = 450: .Enabled = f!uxNoteDrafts.ListCount > 0: End With
    f.Section(acDetail).Height = 6300
    UXNotesSection = True
End Function
Public Function UXSelectReportNote() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmEditor!uxFieldSurface.Form
    f!uxNoteDetails = DLookup("NoteText", "ReportNote", "NoteID=" & Nz(f!uxNoteList, 0) & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0))
    UXSelectReportNote = True
End Function
Public Function UXOpenNote(Optional ByVal recover As Boolean = False) As Boolean
    Dim guid As String
    If CurrentProject.AllForms("frmUXNoteComposer").IsLoaded Then Forms!frmUXNoteComposer.SetFocus: Exit Function
    If recover Then guid = Nz(Forms!frmEditor!uxFieldSurface.Form!uxNoteDrafts, ""): If guid = "" Then Exit Function
    DoCmd.OpenForm "frmUXNoteComposer", , , , , , Nz(Forms!frmEditor!txtReportID, 0) & ":" & guid
    UXOpenNote = True
End Function
Public Function UXLoadNoteComposer() As Boolean
    Dim args As Variant, r As DAO.Recordset, f As Access.Form
    Set f = Forms!frmUXNoteComposer: args = Split(Nz(f.OpenArgs, "0:"), ":")
    noteReport = Val(args(0)): noteGuid = args(1): noteVersion = 0: noteAuthor = MajorAuthor(): noteFinished = False
    If noteGuid = "" Then
        noteGuid = NewIdentity(): f!txtNote = Null
    Else
        Set r = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(noteGuid), dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 6954, , "Note checkpoint no longer exists."
        If r!ReportID <> noteReport Or r!AuthorProfileID <> noteAuthor Then Err.Raise vbObjectError + 6954, , "This checkpoint belongs to another contributor or report."
        noteVersion = r!VersionCounter: f!txtNote = r!NoteText: r.Close
    End If
    f!lblContributor.Caption = MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & noteAuthor)
    f!lblStatus.Caption = "Local note for " & UXMemberLabel(noteReport) & ". Post records an attributed entry; it does not save report edits."
    UXLoadNoteComposer = True
End Function
Private Function NotePendingText() As String
    NotePendingText = Nz(Forms!frmUXNoteComposer!txtNote, "")
    On Error Resume Next
    If Screen.ActiveControl.Name = "txtNote" And Screen.ActiveForm.Name = "frmUXNoteComposer" Then NotePendingText = Forms!frmUXNoteComposer!txtNote.Text
End Function
Public Function UXNoteCheckpoint() As Boolean
    Dim text As String
    On Error GoTo Failed
    If noteFinished Then UXNoteCheckpoint = True: Exit Function
    If MajorAuthor() <> noteAuthor Then Err.Raise vbObjectError + 6954, , "Contributor changed. Return to the original contributor before posting."
    text = NotePendingText()
    If Len(text) = 0 And noteVersion = 0 Then UXNoteCheckpoint = True: Exit Function
    noteVersion = SaveReportNoteDraft(noteGuid, noteReport, noteVersion, text)
    Forms!frmUXNoteComposer!lblStatus.Caption = "Unposted note checkpoint saved at " & Format$(Now(), "hh:nn:ss") & ". Post when ready."
    UXNoteCheckpoint = True: Exit Function
Failed:
    Forms!frmUXNoteComposer!lblStatus.Caption = Err.Description & " Keep this window open, retry, or Discard this note."
End Function
Public Function UXPostNote() As Boolean
    Dim id As Long
    On Error GoTo Failed
    If Not UXNoteCheckpoint() Then Exit Function
    id = PostReportNote(noteGuid, noteVersion): noteFinished = True
    DoCmd.Close acForm, "frmUXNoteComposer", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then UXSelectSection 7
    UXPostNote = True: Exit Function
Failed:
    Forms!frmUXNoteComposer!lblStatus.Caption = Err.Description
End Function
Public Function UXDiscardNote() As Boolean
    'Never delete another window's newer checkpoint when explicitly discarding.
    On Error Resume Next
    If noteVersion > 0 Then DiscardReportNoteDraft noteGuid, noteVersion
    On Error GoTo 0
    noteFinished = True: DoCmd.Close acForm, "frmUXNoteComposer", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then UXSelectSection 7
    UXDiscardNote = True
End Function
Public Function UXNoteUnload() As Boolean
    If Not noteFinished Then
        If Not UXNoteCheckpoint() Then DoCmd.CancelEvent: Exit Function
    End If
    UXNoteUnload = True
End Function
Public Function UXNoteClosed() As Boolean
    If CurrentProject.AllForms("frmEditor").IsLoaded Then UXSelectSection 7
    UXNoteClosed = True
End Function

Public Function UXKeepNoteCheckpoint() As Boolean
    If Not UXNoteCheckpoint() Then Exit Function
    noteFinished = True: DoCmd.Close acForm, "frmUXNoteComposer", acSaveNo
    UXKeepNoteCheckpoint = True
End Function
Public Function UXOpenNoteRecoveries() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmUXNoteRecovery"
    Forms!frmUXNoteRecovery!lstDrafts.RowSource = "SELECT DraftGuid,UXMemberLabel(ReportID) AS MemberName,Format(CheckpointAt,'yyyy-mm-dd hh:nn') AS Checkpoint FROM ReportNoteDraft WHERE AuthorProfileID=" & MajorAuthor() & " ORDER BY CheckpointAt DESC"
    UXOpenNoteRecoveries = True
End Function
Public Function UXResumeSelectedNote() As Boolean
    Dim guid As String, r As DAO.Recordset, id As Long
    On Error GoTo Failed
    guid = Nz(Forms!frmUXNoteRecovery!lstDrafts, "")
    Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(guid) & " AND AuthorProfileID=" & MajorAuthor(), dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 6956, , "Select one of your unposted note checkpoints."
    id = r!ReportID: r.Close
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): UXSelectSection 7
    DoCmd.OpenForm "frmUXNoteComposer", , , , , , CStr(id) & ":" & guid
    DoCmd.Close acForm, "frmUXNoteRecovery", acSaveNo
    UXResumeSelectedNote = True: Exit Function
Failed:
    Forms!frmUXNoteRecovery!lblStatus.Caption = Err.Description
End Function
