Attribute VB_Name = "AuditConversationHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 9090, , message
 n = n + 1
End Sub
Private Function RejectSave() As Boolean
 On Error GoTo Rejected
 CommitDraft 0, "EVAL", "Must roll back", "Synthetic", Null, Null, ""
 Exit Function
Rejected:
 RejectSave = True
End Function
Private Function RejectLegacyState(ByVal id As Long) As Boolean
 On Error GoTo Rejected
 SetReviewThreadState id, True
 Exit Function
Rejected:
 RejectLegacyState = True
End Function
Private Function RejectConversion() As Boolean
 On Error GoTo Rejected
 MigrateReportMessages
 Exit Function
Rejected:
 RejectConversion = True
End Function
Private Function RejectPrivatePost(ByVal guid As String, ByVal version As Long) As Boolean
 On Error GoTo Rejected
 PostReportMessageDraft guid, version
 Exit Function
Rejected:
 RejectPrivatePost = True
End Function
Private Sub LegacyComment(ByVal id As Long, ByVal report As Long, ByVal revision As Long, ByVal text As String)
 Dim r As DAO.Recordset
 Set r = CurrentDb.OpenRecordset("ReviewComment", dbOpenDynaset)
 r.AddNew: r!CommentID = id: r!CommentGuid = NewIdentity(): r!ReportID = report: r!RevisionID = revision
 r!AuthorProfileID = MajorAuthor(): r!CommentAt = #1/2/2024#: r!CommentBody = text: r.Update: r.Close
End Sub
Public Function TestAuditConversation() As String
 Dim report As Long, other As Long, rev As Long, op As Long, folder As Long, batch As Long, root As Long, reply As Long, note As Long
 Dim guid As String, version As Long, legacyCount As Long, eventCount As Long, before As Long, beforeRev As Long, r As DAO.Recordset, actor As Long, firstGuid As String, secondActor As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC AUDIT AND CONVERSATIONS", "Audit actor"
 actor = MajorAuthor()
 report = CommitDraft(0, "EVAL", "Report", "Synthetic", Null, Null, "")
 other = CommitDraft(0, "FITREP", "Other", "Synthetic", Null, Null, "")
 rev = CurrentRevision(report)
 Check PersistenceCount("AuditEvent", "RevisionSourceID=" & rev) = 1, "Revision lacks one canonical event"
 Check DLookup("ActorID", "AuditEvent", "RevisionSourceID=" & rev) = actor, "Revision actor changed"
 Check DLookup("EventAt", "AuditEvent", "RevisionSourceID=" & rev) = DLookup("CommittedAt", "ReportRevision", "RevisionID=" & rev), "Original revision time changed"
 firstGuid = DLookup("AuditGuid", "AuditEvent", "RevisionSourceID=" & rev)
 op = BeginOperation("Synthetic operation", "Workspace operation")
 Set r = CurrentDb.OpenRecordset("UserProfile", dbOpenDynaset)
 r.AddNew: r!ProfileGuid = NewIdentity(): r!DisplayName = "Second actor": r.Update: r.Bookmark = r.LastModified: secondActor = r!UserProfileID: r.Close
 SwitchContributor secondActor
 FinishOperation op, "Prepared", "First outcome"
 Check DLookup("ActorID", "AuditEvent", "JournalSourceID=" & op & " AND SourceVersion=1") = secondActor, "Finishing actor was attributed to the initiator"
 Check DLookup("AuthorProfileID", "OperationJournal", "OperationID=" & op) = actor, "Original operation author changed"
 SwitchContributor actor
 FinishOperation op, "Completed", "Second outcome"
 Check PersistenceCount("AuditEvent", "JournalSourceID=" & op) = 3, "Operation start/outcomes were overwritten or duplicated"
 Check DLookup("EventDetail", "AuditEvent", "JournalSourceID=" & op & " AND SourceVersion=1") = "First outcome", "Earlier operation outcome was lost"
 Check DLookup("CorrelationGuid", "AuditEvent", "JournalSourceID=" & op) = DLookup("OperationGuid", "OperationJournal", "OperationID=" & op), "Operation correlation differs"
 ArchiveDraft report, True, op: ArchiveDraft report, False, op
 Check PersistenceCount("ReportEvent", "ReportID=" & report) = 2, "Distinct archive/recovery history was removed"
 Check PersistenceCount("AuditEvent", "SourceKind='Report event' AND ReportID=" & report & " AND OperationID=" & op) = 2, "Report event context/correlation missing"
 Check CurrentRevision(report) = rev, "Archive/recovery changed content revision"
 folder = CreateFolder("Synthetic folder")
 batch = ApplyFolderBatch(CStr(report), folder, FolderBatchStamp(CStr(report)))
 Check PersistenceCount("AuditEvent", "SourceKind='Folder move' AND ReportID=" & report & " AND OperationID=" & batch) = 1, "Folder move correlation missing"
 Check UndoFolderBatch(batch), "Folder undo failed"
 Check IsNull(DLookup("FolderID", "Report", "ReportID=" & report)), "Folder undo changed final folder"
 Check PersistenceCount("FolderMove", "ReportID=" & report) = 2, "Compensating folder evidence lost"
 guid = NewIdentity(): version = SaveReportNoteDraft(guid, report, 0, "Canonical note via old API")
 note = PostReportNote(guid, version)
 Check note < 0 And PersistenceCount("ReportNote", "ReportID=" & report) = 0, "Old note API created a legacy row"
 root = AddReviewComment(report, rev, "Canonical conversation via compatibility API")
 reply = AddReviewComment(report, rev, "Canonical reply", root)
 Check root < 0 And reply < 0 And DLookup("ParentMessageID", "ReportMessage", "MessageID=" & -reply) = -root, "Canonical token reply lost its parent"
 Check PersistenceCount("ReviewComment", "ReportID=" & report) = 0, "Comment API still writes legacy comments"
 Check RejectLegacyState(root), "Retired resolution API writes new state"
 Check PersistenceCount("ReviewCommentEvent", "1=1") = 0, "Retired resolution wrote legacy events"
 CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & report, dbFailOnError
 reply = AddReportMessage(report, "Notes remain independent of approval", -root)
 Check reply > 0 And CurrentRevision(report) = rev, "Approval blocked notes or changed revision"
 CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & report, dbFailOnError
 guid = NewIdentity(): version = SaveReportMessageDraft(guid, report, 0, "Private pending reply", -root)
 Check PersistenceCount("ReportNoteDraft", "DraftGuid=" & QuoteSQL(guid) & " AND AuthorProfileID=" & actor) = 1, "Private checkpoint identity lost"
 SwitchContributor secondActor
 Check RejectPrivatePost(guid, version), "Another contributor posted a private draft"
 SwitchContributor actor
 LegacyComment 601, report, rev, "Child before parent"
 LegacyComment 602, report, rev, "Nested child"
 LegacyComment 603, report, rev, "Later-numbered root"
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=603 WHERE CommentID=601", dbFailOnError
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=601 WHERE CommentID=602", dbFailOnError
 Set r = CurrentDb.OpenRecordset("ReportNote", dbOpenDynaset)
 r.AddNew: r!NoteGuid = NewIdentity(): r!ReportID = report: r!ContributorName = String$(160, "N"): r!NoteText = Null: r!CreatedAt = #1/2/2024#: r!SourceKind = "Legacy": r.Update: r.Close
 legacyCount = DCount("*", "ReviewComment") + DCount("*", "ReportNote")
 Check MigrateReportMessages() = 4, "Legacy conversion omitted rows"
 root = DLookup("MessageID", "ReportMessage", "LegacyCommentID=603")
 reply = DLookup("MessageID", "ReportMessage", "LegacyCommentID=602")
 Check ReportMessageRoot(reply, report) = root, "Out-of-order nested history lost its root"
 Check Len(DLookup("ContributorName", "ReportMessage", "LegacyNoteID Is Not Null")) = 160, "Legacy contributor was truncated"
 Check IsNull(DLookup("MessageBody", "ReportMessage", "LegacyNoteID Is Not Null")), "Legacy Null body was coerced"
 before = DCount("*", "ReportMessage"): eventCount = DCount("*", "AuditEvent")
 Check MigrateReportMessages() = 0, "Conversation conversion is not idempotent"
 Call MigrateAuditHistory: MigrateAuditHistory
 Check DCount("*", "ReportMessage") = before, "Audit backfill changed conversations"
 Check DLookup("AuditGuid", "AuditEvent", "RevisionSourceID=" & rev) = firstGuid, "Audit backfill replaced existing identity"
 Check DCount("*", "ReviewComment") + DCount("*", "ReportNote") = legacyCount, "Legacy evidence rows were changed"
 LegacyComment 701, other, CurrentRevision(other), "Invalid cross-report child"
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=603 WHERE CommentID=701", dbFailOnError
 Check RejectConversion() And DCount("*", "ReportMessage") = before, "Cross-report parent was accepted or partially copied"
 CurrentDb.Execute "DELETE FROM ReviewComment WHERE CommentID=701", dbFailOnError
 LegacyComment 801, report, rev, "Cycle A": LegacyComment 802, report, rev, "Cycle B"
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=802 WHERE CommentID=801", dbFailOnError
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=801 WHERE CommentID=802", dbFailOnError
 Check RejectConversion() And DCount("*", "ReportMessage") = before, "Cycle was accepted or partially copied"
 CurrentDb.Execute "UPDATE ReviewComment SET ThreadID=Null WHERE CommentID In (801,802)", dbFailOnError
 CurrentDb.Execute "DELETE FROM ReviewComment WHERE CommentID In (801,802)", dbFailOnError
 before = DCount("*", "Report"): beforeRev = DCount("*", "ReportRevision"): eventCount = DCount("*", "AuditEvent")
 CurrentDb.TableDefs("AuditEvent").Fields("EventKind").ValidationRule = "='Reject audit'"
 Check RejectSave(), "Injected audit failure was accepted"
 CurrentDb.TableDefs("AuditEvent").Fields("EventKind").ValidationRule = ""
 Check DCount("*", "Report") = before And DCount("*", "ReportRevision") = beforeRev And DCount("*", "AuditEvent") = eventCount, "Audit failure did not roll back domain change/envelope"
 TestAuditConversation = "PASS: " & n & " audit identity/correlation, retained history, canonical messages, legacy conversion and rollback assertions": Exit Function
Failed:
 TestAuditConversation = "FAIL after " & n & ": " & Err.Source & ": " & Err.Number & " " & Err.Description
End Function

Public Function TestNavigation() As String
 Dim report As Long, rev As Long, row As Variant, parts As Variant, target As Variant, before As String
 On Error GoTo Failed
 report = DMin("ReportID", "Report"): rev = CurrentRevision(report)
 DoCmd.OpenForm "frmEditor", , , , , , CStr(report)
 Forms!frmEditor!txtTitle = "Pending alias value"
 For Each target In Array("frmReviewComments", "frmUXNoteComposer", "frmUXReview", "frmChecklist")
  Check OpenScreen(CStr(target)), "Compatibility tool route failed"
  Check Forms!frmEditor!txtTitle = "Pending alias value" And CurrentRevision(report) = rev, "Tool route saved/lost pending edits"
 Next
 Check OpenScreen("frmHelp"), "Help alias failed"
 Check Forms!frmEditor!txtTitle = "Pending alias value", "Help changed report buffer"
 DoCmd.Close acForm, "frmUXHelp", acSaveNo
 Forms!frmEditor!txtStart = "Invalid pending date"
 Check UXNavigate("frmLibrary"), "Checkpointed invalid draft could not navigate"
 Check Not CurrentProject.AllForms("frmEditor").IsLoaded And CurrentRevision(report) = rev, "Invalid navigation committed a revision or kept editor open"
 Check DLookup("StartEntry", "RecoveryDraft", "ReportID=" & report) = "Invalid pending date", "Navigation lost invalid raw date"
 Check RestoreRecoveryDraft(DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & report)), "Navigation checkpoint could not recover"
 Forms!frmEditor!txtStart = Null
 Check UXNavigate("frmLibrary"), "Valid navigation failed"
 Check DLookup("Title", "Report", "ReportID=" & report) = "Pending alias value", "Navigation omitted valid save"
 For Each row In UXWorkspaceRoutes()
  parts = Split(CStr(row), "|")
  Check UXNavigate(CStr(parts(1))), "Workspace destination failed: " & parts(1)
 Next
 Check UXOpenReportTool(report, "Notes"), "Return to report conversation failed"
 UXToolsForm()!lstMessages = DLookup("MessageID", "ReportMessage", "LegacyCommentID=603")
 Check UXMessageRead(), "Nested conversation display failed"
 Check InStr(Nz(UXToolsForm()!txtConversation, ""), "Nested child") > 0, "Nested legacy reply is invisible"
 Check UXOpenValidation(False), "Validation popout failed"
 UXLeaveEditor "Discard"
 TestNavigation = "PASS: " & n & " total audit/conversation/navigation assertions; canonical destinations, aliases and pending guards": Exit Function
Failed:
 TestNavigation = "FAIL after " & n & ": " & Err.Source & ": " & Err.Number & " " & Err.Description
End Function
