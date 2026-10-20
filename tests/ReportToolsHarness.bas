Attribute VB_Name = "ReportToolsHarness"
Option Explicit
Private count As Long, stage As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise 5, , message
 count = count + 1
End Sub
Public Function TestReportTools() As String
 Dim id As Long, rev As Long, root As Long, reply As Long, legacy As Long, before As Long, f As Access.Form, p As Access.Form, v As Variant, guid As String
 On Error GoTo Failed
 stage = "legacy preservation"
 id = DLookup("ReportID", "Report", "WorkflowState='Approved'"): rev = CurrentRevision(id)
 legacy = AddReviewComment(id, rev, "Historical conversation")
 SetReviewThreadState legacy, True
 before = DCount("*", "ReviewCommentEvent")
 MigrateReportMessages
 root = DLookup("MessageID", "ReportMessage", "LegacyCommentID=" & legacy)
 reply = AddReportMessage(id, "Reply after approval and historical resolution", root)
 Check DLookup("ParentMessageID", "ReportMessage", "MessageID=" & reply) = root, "Reply parent lost"
 Check DCount("*", "ReviewCommentEvent") = before, "Historical events changed"
 Check MigrateReportMessages() = 0, "Conversion duplicated legacy records"
 Check CurrentRevision(id) = rev, "Communication changed report revision"
 stage = "embedded notes"
 Check UXOpenReportTool(id, "Notes", 0, 0, 0, root), "Open notes failed"
 Set f = Forms!frmEditor: Set p = UXToolsForm()
 Check p!txtView = "Notes" And p!txtMessage.Visible, "Notes view not embedded"
 Check Len(Nz(f!uxError.Caption, "")) = 0, "Notes error: " & f!uxError.Caption
 Check InStr(p!txtConversation, "Reply after approval") > 0, "Conversation missing reply"
 Check UXMessageBegin(True), "Begin reply failed"
 p!txtMessage = "Private reply checkpoint": Check UXMessageKeep(), "Save private reply failed"
 guid = p!txtDraftGuid
 Check DLookup("ParentMessageID", "ReportNoteDraft", "DraftGuid=" & QuoteSQL(guid)) = root, "Draft reply parent lost"
 For Each v In Array("Validation", "Calculations", "Review", "Notes")
  stage = "view " & v
  f!uxError.Caption = ""
  Check UXToolsChoose(CStr(v)), "View switch failed"
  Check p!txtView = v, "Wrong embedded view"
  Check Len(Nz(f!uxError.Caption, "")) = 0, "View error: " & f!uxError.Caption
 Next
 Check p!txtMessage = "Private reply checkpoint", "Tab switch lost unsent text"
 Check UXMessagePost(), "Posting failed: " & p!lblNoteStatus.Caption
 Check DCount("*", "ReportNoteDraft", "DraftGuid=" & QuoteSQL(guid)) = 0, "Posted checkpoint retained"
 stage = "validation popup"
 Check UXOpenValidation(), "Popup open failed"
 Check Forms!frmUXValidation.PopUp And Not Forms!frmUXValidation.Modal, "Validation is not modeless popup"
 Check CurrentProject.AllForms("frmEditor").IsLoaded, "Popup replaced editor"
 stage = "return navigation"
 Check UXWorkBack(), "Return to My Work failed"
 Check Screen.ActiveForm.Name = "frmMyWork", "My Work did not receive focus"
 Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Editor left open"
 Check Not CurrentProject.AllForms("frmUXValidation").IsLoaded, "Validation orphaned"
 Check UXMyWorkMode("Notes"), "My Work Notes failed"
 Check Forms!frmMyWork!lstReportNotes.Visible, "My Work notes not visible"
 TestReportTools = "PASS: " & count & " report-tools checks": Exit Function
Failed:
 TestReportTools = "FAIL " & stage & ": " & Err.Number & " " & Err.Description
End Function
