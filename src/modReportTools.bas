Attribute VB_Name = "modReportTools"
Option Explicit
Private selectedView As String, findings As Collection, draftTick As Date
Public Function UXToolsForm() As Access.Form
 Set UXToolsForm = Forms!frmEditor!uxToolsSurface.Form
End Function
Public Function UXReviewForm() As Access.Form
 Set UXReviewForm = UXToolsForm()!uxReviewPanel.Form
End Function
Public Function UXToolsView() As String
 UXToolsView = selectedView
End Function
Public Function UXToolsChoose(ByVal view As String) As Boolean
 selectedView = view
 UXToolsChoose = UXSelectSection(8)
End Function
Public Function UXToolsAvailable() As Boolean
 If CurrentProject.AllForms("frmEditor").IsLoaded Then UXToolsAvailable = UXHasControl(Forms!frmEditor, "uxToolsSurface")
End Function
Public Function UXToolsRefresh() As Boolean
 Dim f As Access.Form, c As Access.Control, id As Long, name As Variant, active As Boolean
 On Error GoTo Failed
 Set f = UXToolsForm(): id = Nz(Forms!frmEditor!txtReportID, 0)
 If Nz(f!txtReportID, 0) <> id Then
  f!txtReportID = id: f!txtDraftGuid = NewIdentity(): f!txtDraftVersion = 0: f!txtDraftParent = 0: f!txtDraftSaved = Null: f!txtMessage = Null
 End If
 If selectedView = "" Then selectedView = "Notes"
 f!uxToolFocus.SetFocus
 f!txtView = selectedView
 For Each c In f.Controls
  If Left$(c.Tag, 5) = "VIEW:" Then c.Visible = (Mid$(c.Tag, 6) = selectedView)
 Next
 f!uxReviewPanel.Visible = (selectedView = "Review")
 For Each name In Array("cmdComments", "cmdRoute", "cmdChecklist", "cmdCalculations")
  With Forms!frmEditor.Controls(name)
   active = (CStr(name) = "cmdComments" And selectedView = "Notes") Or (CStr(name) = "cmdRoute" And selectedView = "Review") Or (CStr(name) = "cmdChecklist" And selectedView = "Validation") Or (CStr(name) = "cmdCalculations" And selectedView = "Calculations")
   .FontBold = active: .BackColor = IIf(active, UXColor("Primary"), UXColor("Neutral"))
   .ForeColor = IIf(active, UXColor("White"), UXColor("Ink"))
  End With
 Next
 Select Case selectedView
  Case "Notes": UXMessageRefresh
  Case "Review"
   If id > 0 Then UXLoadReview id Else f!lblToolStatus.Caption = "Save the report before assigning saved-revision review."
   UXReportReviewMode
  Case "Validation": UXToolsValidation
  Case "Calculations": UXToolsCalculations
 End Select
 UXToolsRefresh = True: Exit Function
Failed:
 Forms!frmEditor!uxError.Caption = "Report tools: " & Err.Description
End Function
Private Function MessageText() As String
 Dim f As Access.Form
 Set f = UXToolsForm(): MessageText = Nz(f!txtMessage, "")
 On Error Resume Next
 If Screen.ActiveControl.Name = "txtMessage" Then MessageText = f!txtMessage.Text
 On Error GoTo 0
End Function
Public Function UXToolsCheckpoint(Optional ByVal closing As Boolean = False) As Boolean
 Dim f As Access.Form, text As String, r As Access.Form
 On Error GoTo Failed
 If Not UXToolsAvailable() Then UXToolsCheckpoint = True: Exit Function
 Set f = UXToolsForm()
 If closing Then
  Set r = UXReviewForm()
  If Len(Nz(r!txtNote, "")) > 0 Or Len(Nz(r!txtAssignmentNote, "")) > 0 Then Err.Raise 5, "modReportTools.UXToolsCheckpoint", "Record or clear the unfinished review decision note before leaving this report."
 End If
 'Messages post directly; there are no private drafts. Leaving with typed text asks
 'for it to be posted or cleared. A legacy draft resumed from recovery keeps saving.
 text = MessageText()
 If Nz(f!txtDraftVersion, 0) > 0 And Nz(f!txtReportID, 0) > 0 Then
  If text <> Nz(f!txtDraftSaved, "") Then f!txtDraftVersion = SaveReportMessageDraft(f!txtDraftGuid, f!txtReportID, Nz(f!txtDraftVersion, 0), text, Nz(f!txtDraftParent, 0)): f!txtDraftSaved = text
 ElseIf closing And Len(Trim$(text)) > 0 Then
  Err.Raise 5, "modReportTools.UXToolsCheckpoint", "Post or clear your message before leaving this report."
 End If
 draftTick = Now: UXToolsCheckpoint = True: Exit Function
Failed:
 Forms!frmEditor!uxError.Caption = "Note not saved: " & Err.Description
End Function
Public Function UXToolsTimer() As Boolean
 If Not UXToolsAvailable() Then Exit Function
 If DateDiff("s", draftTick, Now) >= 30 Then UXToolsCheckpoint
 UXToolsTimer = True
End Function
Public Function UXMessageRefresh() As Boolean
 Dim f As Access.Form, id As Long, old As Variant, editable As Boolean
 Set f = UXToolsForm(): id = Nz(f!txtReportID, 0): old = f!lstMessages
 f!lstMessages.RowSource = "SELECT MessageID,Left(MessageBody,110) AS [Conversation],ContributorName AS Contributor,CreatedAt AS Posted FROM ReportMessage WHERE ReportID=" & id & " AND ParentMessageID Is Null ORDER BY CreatedAt DESC,MessageID DESC"
 f!lstMessages.Requery
 If Not IsNull(old) Then f!lstMessages = old
 If IsNull(f!lstMessages) And f!lstMessages.ListCount > 1 Then f!lstMessages = f!lstMessages.ItemData(1)
 f!lstDrafts.RowSource = "SELECT DraftGuid,Left(NoteText,80) AS [Your unfinished notes],CheckpointAt FROM ReportNoteDraft WHERE ReportID=" & id & " AND AuthorProfileID=" & MajorAuthor() & " ORDER BY CheckpointAt DESC"
 f!lstDrafts.Requery
 editable = id > 0 And Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True)
 f!txtMessage.Locked = Not editable
 f!cmdNew.Enabled = editable: f!cmdReply.Enabled = editable And Not IsNull(f!lstMessages): f!cmdPost.Enabled = editable
 f!cmdKeepDraft.Enabled = editable: f!cmdDiscardDraft.Enabled = editable: f!cmdResume.Enabled = editable
 f!lblCompose.Caption = "New message as " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & MajorAuthor()), "the current contributor")
 f!lblNoteStatus.Caption = "Messages stay with this report, independent of review or approval, and can continue after it is signed."
 If id = 0 Then f!lblNoteStatus.Caption = "Save this report once before posting notes."
 UXMessageRead: UXMessageRefresh = True
End Function
Public Function UXMessageRead() As Boolean
 Dim f As Access.Form, rows As DAO.Recordset, id As Long, text As String, parents As Object
 'One running log, oldest first: every message and reply with its author and time.
 Set f = UXToolsForm()
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReportMessage WHERE ReportID=" & Nz(f!txtReportID, 0) & " ORDER BY CreatedAt,MessageID", dbOpenSnapshot)
 Do While Not rows.EOF
  text = text & rows!ContributorName & "  " & Nz(Format(rows!CreatedAt, "yyyy-mm-dd hh:nn"), "") & vbCrLf & Nz(rows!MessageBody, "") & vbCrLf & vbCrLf
  rows.MoveNext
 Loop
 rows.Close
 If text = "" Then text = "No messages yet."
 f!txtConversation = text: UXMessageRead = True
End Function
Public Function UXMessageBegin(ByVal reply As Boolean) As Boolean
 Dim f As Access.Form, parent As Long
 Set f = UXToolsForm()
 If reply Then parent = Nz(f!lstMessages, 0): If parent = 0 Then Exit Function
 If Not UXToolsCheckpoint() Then Exit Function
 f!txtDraftGuid = NewIdentity(): f!txtDraftVersion = 0: f!txtDraftParent = parent: f!txtDraftSaved = Null: f!txtMessage = Null
 f!lblCompose.Caption = IIf(reply, "Reply to selected conversation", "Write a new report note")
 f!txtMessage.SetFocus: UXMessageBegin = True
End Function
Public Function UXMessagePost() As Boolean
 Dim f As Access.Form, id As Long
 On Error GoTo Failed
 Set f = UXToolsForm()
 If Len(Trim$(MessageText())) = 0 Then f!lblNoteStatus.Caption = "Type a message first.": Exit Function
 If Nz(f!txtDraftVersion, 0) > 0 Then
  If Not UXToolsCheckpoint() Then Exit Function
  id = PostReportMessageDraft(Nz(f!txtDraftGuid, ""), Nz(f!txtDraftVersion, 0))
 Else
  id = AddReportMessage(Nz(f!txtReportID, 0), MessageText(), Nz(f!txtDraftParent, 0))
 End If
 f!txtMessage = Null: f!txtDraftSaved = Null: f!txtDraftVersion = 0: f!txtDraftParent = 0: f!txtDraftGuid = NewIdentity()
 If Not IsNull(DLookup("ParentMessageID", "ReportMessage", "MessageID=" & id)) Then id = DLookup("ParentMessageID", "ReportMessage", "MessageID=" & id)
 f!lstMessages = id: UXMessageRefresh
 UXMessagePost = True: Exit Function
Failed:
 f!lblNoteStatus.Caption = Err.Description
End Function
Public Function UXMessageKeep() As Boolean
 If Not UXToolsCheckpoint() Then Exit Function
 UXMessageRefresh: UXToolsForm()!lblNoteStatus.Caption = "Your unfinished note is saved privately. It has not been posted."
 UXMessageKeep = True
End Function
Public Function UXMessageDiscard() As Boolean
 Dim f As Access.Form
 On Error GoTo Failed
 Set f = UXToolsForm()
 If Nz(f!txtDraftVersion, 0) > 0 Then DiscardReportNoteDraft f!txtDraftGuid, f!txtDraftVersion
 f!txtMessage = Null: f!txtDraftSaved = Null: f!txtDraftGuid = NewIdentity(): f!txtDraftVersion = 0: f!txtDraftParent = 0
 UXMessageRefresh: UXMessageDiscard = True: Exit Function
Failed:
 f!lblNoteStatus.Caption = Err.Description
End Function
Public Function UXMessageResume(Optional ByVal identity As String = "") As Boolean
 Dim f As Access.Form, rows As DAO.Recordset
 On Error GoTo Failed
 Set f = UXToolsForm()
 If identity = "" Then identity = Nz(f!lstDrafts, "")
 If identity = "" Then Exit Function
 If Not UXToolsCheckpoint() Then Exit Function
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(identity) & " AND AuthorProfileID=" & MajorAuthor() & " AND ReportID=" & f!txtReportID, dbOpenSnapshot)
 If rows.EOF Then Err.Raise 5, "modReportTools.UXMessageResume", "Choose one of your unfinished notes for this report."
 f!txtDraftGuid = rows!DraftGuid: f!txtDraftVersion = rows!VersionCounter: f!txtDraftParent = Nz(rows!ParentMessageID, 0)
 f!txtMessage = rows!NoteText: f!txtDraftSaved = rows!NoteText: rows.Close
 f!lblCompose.Caption = IIf(Nz(f!txtDraftParent, 0) > 0, "Continue your draft reply", "Continue your draft note")
 UXMessageResume = True: Exit Function
Failed:
 f!lblNoteStatus.Caption = Err.Description
End Function
Public Function UXToolsValidation() As Boolean
 Dim f As Access.Form, item As Object, i As Long, values As Object
 Set f = UXToolsForm(): Set values = EditorPendingFields(): Set findings = ValidateLiveFields(values, Nz(Forms!frmEditor!txtReportID, 0))
 f!lstFindings.RowSource = "": f!lstFindings.RowSourceType = "Value List"
 For Each item In findings
  i = i + 1: f!lstFindings.AddItem i & ";" & FindingColumn(item("Label")) & ";" & FindingColumn(item("Severity")) & ";" & FindingColumn(Left$(item("Message"), 140))
 Next
 f!lblValidation.Caption = i & " findings - includes pending report edits. Save remains available for incomplete drafts."
 f!txtEvidence = "No findings from implemented checks."
 If UXHasControl(f, "cmdApplyCounseling") Then f!cmdApplyCounseling.Enabled = False
 If i > 0 Then f!lstFindings = 0 + f!lstFindings.ItemData(0): UXToolsFinding
 UXToolsValidation = True
End Function
'Value lists also split on commas, so each column is quoted (labels such as
'"Name (Last, First MI Suffix)" and messages contain commas).
Private Function FindingColumn(ByVal value As Variant) As String
 FindingColumn = """" & Replace(Replace(Nz(value, ""), ";", ","), """", "'") & """"
End Function
Public Function UXToolsFinding() As Boolean
 Dim item As Object
 If findings Is Nothing Then Exit Function
 If IsNull(UXToolsForm()!lstFindings) Then Exit Function
 Set item = findings(CLng(UXToolsForm()!lstFindings))
 UXToolsForm()!txtEvidence = item("Severity") & ": " & item("Message") & vbCrLf & vbCrLf & item("Reference")
 If UXHasControl(UXToolsForm(), "cmdApplyCounseling") Then UXToolsForm()!cmdApplyCounseling.Enabled = (item("Code") = "COUNSEL-RECORD")
 UXToolsFinding = True
End Function
Public Function UXToolsFindingOpen() As Boolean
 Dim item As Object
 If findings Is Nothing Then Exit Function
 If IsNull(UXToolsForm()!lstFindings) Then Exit Function
 Set item = findings(CLng(UXToolsForm()!lstFindings))
 UXToolsFindingOpen = UXFixedOpenField(CStr(item("Key")))
End Function
Public Function UXToolsCalculations() As Boolean
 Dim text As String, id As Long, group As Long
 id = Nz(Forms!frmEditor!txtReportID, 0): group = Nz(Forms!frmEditor!uxGroup, 0)
 text = "Current editor values (including pending edits):" & vbCrLf & UXTraitContext(EditorPendingFields())
 If id > 0 Then text = text & vbCrLf & vbCrLf & "Latest saved revision:" & vbCrLf & UXTraitContext(SavedFields(id))
 If group > 0 Then text = text & vbCrLf & vbCrLf & UXGroupCalculationText(group)
 UXToolsForm()!txtCalculations = text: UXToolsCalculations = True
End Function
Public Function UXOpenReportTool(ByVal report As Long, ByVal view As String, Optional ByVal group As Long = 0, Optional ByVal run As Long = 0, Optional ByVal assignment As Long = 0, Optional ByVal message As Long = 0) As Boolean
 On Error GoTo Failed
 If CurrentProject.AllForms("frmEditor").IsLoaded Then
  If Nz(Forms!frmEditor!txtReportID, 0) <> report Then
   If Not UXLeaveEditor() Then Exit Function
  End If
 End If
 If Not UXLeaveGroup() Then Exit Function
 If Not CurrentProject.AllForms("frmEditor").IsLoaded Then DoCmd.OpenForm "frmEditor", , , , , , CStr(report)
 Forms!frmEditor.SetFocus
 If Not UXToolsChoose(view) Then Exit Function
 If view = "Review" Then
  UXLoadReview report, group, run
  UXReviewForm()!cboReviewMode = IIf(assignment > 0, "Assigned reviews", "Route steps"): UXReportReviewMode
  If assignment > 0 Then UXReviewForm()!lstAssignments = assignment: UXReportAssignmentSelect
 ElseIf message > 0 Then
  UXToolsForm()!lstMessages = message: UXMessageRead
 End If
 Forms!frmEditor.SetFocus: UXOpenReportTool = True: Exit Function
Failed:
 If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor!uxError.Caption = Err.Description
End Function
Public Function UXReportReviewMode() As Boolean
 Dim f As Access.Form, c As Access.Control, assigned As Boolean
 Set f = UXReviewForm(): assigned = Nz(f!cboReviewMode, "Route steps") = "Assigned reviews"
 For Each c In f.Controls
  If c.Tag = "ROUTE" Then c.Visible = Not assigned
  If c.Tag = "ASSIGNMENT" Then c.Visible = assigned
 Next
 f!lstAssignments.RowSource = "SELECT t.TaskID,p.DisplayName AS Reviewer,t.TaskState AS Status,t.DueDate AS Due FROM ReviewTask AS t INNER JOIN UserProfile AS p ON t.AssignedTo=p.UserProfileID WHERE t.ReportID=" & Nz(Forms!frmEditor!txtReportID, 0) & " ORDER BY t.TaskID DESC"
 f!lstAssignments.Requery
 If IsNull(f!lstAssignments) And f!lstAssignments.ListCount > 1 Then f!lstAssignments = f!lstAssignments.ItemData(1)
 UXReportAssignmentSelect: UXReportReviewMode = True
End Function
Public Function UXReportAssignmentSelect() As Boolean
 Dim f As Access.Form, rows As DAO.Recordset, id As Long, allowed As Boolean, pid As Long
 Set f = UXReviewForm(): id = Nz(f!lstAssignments, 0): pid = MajorAuthor()
 f!cmdAssignmentAccept.Enabled = False: f!cmdAssignmentReturn.Enabled = False: f!cmdAssignmentCancel.Enabled = False: f!txtAssignmentHistory = "No assignment selected. Assign a saved revision from My Work."
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & id & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0), dbOpenSnapshot)
 If Not rows.EOF Then
  allowed = rows!TaskState = "Pending" And rows!AssignedTo = pid And rows!RevisionID = CurrentRevision(rows!ReportID) And Not Nz(DLookup("IsArchived", "Report", "ReportID=" & rows!ReportID), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & rows!ReportID), "") = "Draft"
  f!cmdAssignmentAccept.Enabled = allowed: f!cmdAssignmentReturn.Enabled = allowed
  f!cmdAssignmentCancel.Enabled = rows!TaskState = "Pending" And rows!RequestedBy = pid
  f!txtAssignmentHistory = rows!TaskState & IIf(rows!RevisionID <> CurrentRevision(rows!ReportID), " / stale saved revision", " / current saved revision") & vbCrLf & Nz(rows!Instructions, "") & vbCrLf & "Requested by " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & rows!RequestedBy), "")
 End If
 rows.Close: UXReportAssignmentSelect = True
End Function
Public Function UXReportAssignmentDecision(ByVal outcome As String) As Boolean
 On Error GoTo Failed
 UXFlushSurface
 If EditorHasPending() And outcome <> "Cancelled" Then Err.Raise 5, "modReportTools.UXReportAssignmentDecision", "Save or undo pending report changes before recording a review decision."
 With UXReviewForm()
  UXReportAssignmentDecision = FinishReview(Nz(!lstAssignments, 0), outcome, Nz(!txtAssignmentNote, ""))
  If UXReportAssignmentDecision Then !txtAssignmentNote = Null
 End With
 UXReportAssignmentSelect: UXWorkChanged: Exit Function
Failed:
 UXReviewForm()!txtAssignmentHistory = Err.Description
End Function
Public Function UXMyWorkNotes() As Boolean
 Dim f As Access.Form
 Set f = Forms!frmMyWork
 f!lstReportNotes.RowSource = "SELECT m.MessageID,m.ReportID,UXWorkReport(m.ReportID) AS Report,m.ContributorName AS Contributor,Left(m.MessageBody,120) AS Note,m.CreatedAt AS Posted FROM ReportMessage AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.ParentMessageID Is Null AND r.IsArchived=False ORDER BY m.CreatedAt DESC,m.MessageID DESC"
 f!lstReportNotes.Requery: UXMyWorkNotes = True
End Function
Public Function UXMyWorkOpenNote() As Boolean
 With Forms!frmMyWork!lstReportNotes
  If IsNull(.Value) Then Exit Function
  UXMyWorkOpenNote = UXOpenReportTool(CLng(.Column(1)), "Notes", 0, 0, 0, CLng(.Column(0)))
 End With
End Function

Public Sub UXToolsSetNotes()
 selectedView = "Notes"
End Sub
'Enter the member's recorded counseling into blocks 30-31 as pending edits. The
'contributor still saves the report; discarding leaves the record unused.
Public Function UXToolsApplyCounseling() As Boolean
 Dim s As Object
 On Error GoTo Failed
 Set s = CounselingSuggestion(EditorPendingFields(), Nz(Forms!frmEditor!txtReportID, 0))
 If s Is Nothing Then UXToolsForm()!lblValidation.Caption = "Blocks 30-31 already match the recorded counseling.": Exit Function
 Call UXSurfaceSetValue("CounselingDate", s("CounselingDate"))
 Call UXSurfaceSetValue("CounselorName", IIf(Nz(s("CounselorName"), "") = "", Null, s("CounselorName")))
 UXToolsValidation
 UXToolsForm()!lblValidation.Caption = "Blocks 30-31 now show the recorded counseling as unsaved edits. Save the report to keep them."
 UXToolsApplyCounseling = True: Exit Function
Failed:
 UXToolsForm()!lblValidation.Caption = "Recorded counseling was not applied: " & Err.Description
End Function