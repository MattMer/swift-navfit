Attribute VB_Name = "CommentsITAHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8990, , message
 n = n + 1
End Sub
Public Function TestCommentsITA() As String
 Dim id As Long, other As Long, rev As Long, root As Long, author As Long, reviewer As Long, note As Long, guid As String, version As Long, c As Access.Control, f As Access.Form, s As Access.Form
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC COMMENTS AND ITA", "First contributor": author = MajorAuthor(): reviewer = AddContributor("Second contributor")
 id = CommitDraft(0, "EVAL", "Comment report", "Synthetic member", Null, Null, "", "Test", "Official narrative unchanged")
 other = CommitDraft(0, "EVAL", "Other report", "Other member", Null, Null, "")
 rev = CurrentRevision(id)
 root = AddReviewComment(id, rev, "Please check the duty description.")
 SwitchContributor reviewer: AddReviewComment id, rev, "Checked the saved duty description.", root
 SwitchContributor author
 Check OpenReviewComments(id), "Open comments failed"
 Set f = Forms!frmReviewComments
 Check Nz(f!lstThreads, 0) = root And InStr(f!txtHistory, "First contributor") > 0 And InStr(f!txtHistory, "Second contributor") > 0, "Conversation not automatically readable"
 Check f!cmdReply.Enabled And f!cmdResolve.Enabled And Not f!cmdReopen.Enabled, "Open comment actions wrong"
 Check UXBeginComment(True), "Reply mode failed"
 Check f!cmdPost.Caption = "Post reply", "Reply intent not labelled"
 f!txtCompose = "Unsent reply"
 Check UXRefreshComments(), "Refresh failed"
 Check f!txtCompose = "Unsent reply" And f!lstThreads = root, "Refresh lost reply or target"
 Check OpenReviewComments(id), "Reopen same report failed"
 Check f!txtCompose = "Unsent reply", "Reopen lost composition"
 Check Not OpenReviewComments(other) And f!txtReportID = id, "Unsent reply moved to another report"
 Check UXSendComment(), "Post reply failed"
 Check DCount("*", "ReviewComment", "ThreadID=" & root) = 2, "Reply not recorded in original conversation"
 Check ReviewStateUI(True), "Resolve failed"
 Check Not f!cmdReply.Enabled And f!cmdReopen.Enabled, "Resolved reply still enabled"
 Check ReviewStateUI(False), "Reopen failed"
 Check UXBeginComment(False), "New comment mode failed"
 f!txtCompose = "A separate comment."
 Check UXSendComment(), "New comment failed"
 Check DCount("*", "ReviewComment", "ReportID=" & id & " AND ThreadID Is Null") = 2, "New conversation not created"
 Check CurrentRevision(id) = rev, "Comments changed official revision"
 CurrentDb.Execute "UPDATE Report SET IsArchived=True WHERE ReportID=" & id, dbFailOnError
 Check OpenReviewComments(id), "Archived comment history inaccessible"
 Check Not f!cmdPost.Enabled And Not f!cmdNewComment.Enabled And f!txtCompose.Locked, "Archived comment form is not read-only"
 CurrentDb.Execute "UPDATE Report SET IsArchived=False WHERE ReportID=" & id, dbFailOnError
 Check DLookup("OfficialComments", "Report", "ReportID=" & id) = "Official narrative unchanged", "Comments altered narrative"
 DoCmd.Close acForm, "frmReviewComments", acSaveNo
 guid = NewIdentity(): version = SaveReportNoteDraft(guid, id, 0, "Working note by another contributor")
 note = PostReportNote(guid, version)
 guid = NewIdentity(): version = SaveReportNoteDraft(guid, id, 0, "Unfinished working note")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Check UXSelectSection(7), "Notes page failed"
 Set s = Forms!frmEditor!uxFieldSurface.Form
 For Each c In s.Controls
  If c.Name Like "uxNotes*" Then Check c.Visible, "Notes label/shortcut absent: " & c.Name
 Next
 Check InStr(s!uxNoteDetails, "Working note") > 0, "Posted note not automatically readable"
 Check s!uxRecoverNote.Enabled And Not IsNull(s!uxNoteDrafts), "Draft continuation unavailable"
 Check s!uxNoteDetails.Locked, "Posted notes editable"
 Check UXOpenNote(True), "Continue draft failed"
 Check Forms!frmUXNoteComposer!txtNote = "Unfinished working note", "Draft body lost"
 Forms!frmUXNoteComposer!txtNote = "Finished working note"
 Check UXPostNote(), "Posting continued note failed"
 Check DLookup("NoteText", "ReportNote", "NoteID=DMax('NoteID','ReportNote')") = "Finished working note", "Posted note body wrong"
 Check DCount("*", "ReportNoteDraft") = 0 And CurrentRevision(id) = rev, "Post lost draft semantics or changed report"
 Check UXSelectSection(8), "Review page failed"
 Check Forms!frmEditor!uxReviewCommentGuide.Visible And Not Forms!frmEditor!uxFieldSurface.Visible, "Review instructions missing or notes retained"
 Check UXLeaveEditor("Discard"), "Editor cleanup failed"
 Check UXITAChange("4.00", "3.50") = 1, "Higher ITA not green category"
 Check UXITAChange("3.00", "3.50") = -1, "Lower ITA not yellow category"
 Check UXITAChange("4.00", "4.00") = 0, "Equal ITA highlighted"
 Check UXITAChange("4.00", Null) = 0 And UXITAChange("4.00", "") = 0, "Missing prior highlighted"
 Check UXITAChange("NOB", "4.00") = 0 And UXITAChange("-", "4.00") = 0 And UXITAChange("4.00", "NOB") = 0, "NOB/non-numeric highlighted"
 Check UXITAChange(Null, Null) = 0 And UXITAChange("4.00", "0") = 0, "Invalid ITA highlighted"
 DoCmd.OpenForm "frmUXTraitRows", acDesign
 Set f = Forms!frmUXTraitRows
 Check f!ITA.FormatConditions.Count = 2, "ITA row conditions absent"
 Check f!ITA.FormatConditions(0).BackColor = UXColor("Success") And f!ITA.FormatConditions(1).BackColor = UXColor("Warning"), "ITA condition colours wrong"
 Check InStr(f!ITA.FormatConditions(0).Expression1, "[PriorITA]") > 0, "ITA colour not based on each row's prior"
 DoCmd.Close acForm, "frmUXTraitRows", acSaveNo
 TestCommentsITA = "PASS: " & n & " notes, conversation intent/preservation and ITA conditional-format assertions": Exit Function
Failed:
 TestCommentsITA = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
