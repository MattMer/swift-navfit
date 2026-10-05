Attribute VB_Name = "UXReviewReturnHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
If Not value Then Err.Raise vbObjectError + 5950, , message
checks = checks + 1
End Sub
Public Function TestReviewReturn() As String
Dim id As Long, template As Long, run As Long, newRun As Long, revision As Long, author As Long, other As Long, i As Long, f As Access.Form, position As Long, topID As Long, failure As Long, deadline As Single, groupID As Long, extraID As Long, c As Access.Control, size As Variant
On Error GoTo Failed
ConfigureWorkspace "SYNTHETIC REVIEW RETURN", "Reviewer"
author = MajorAuthor(): other = AddContributor("Other reviewer")
id = CommitDraft(0, "EVAL", "Review fixture", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
template = CreateRouteTemplate("Named route", CStr(author), "Read all saved blocks")
run = StartRoute(id, template)
DoCmd.OpenForm "frmUXReview", , , , , , CStr(id)
Set f = Forms!frmUXReview
For Each size In Array(9000, 14000, 20000, 9000)
Check UXReviewLayout(CLng(size)), "Review resize failed"
For Each c In f.Controls
If c.Visible Then Check c.Left >= 0 And c.Left + c.Width < CLng(size), "Review clipped horizontally: " & c.Name
Next
Next
Check f!cboRun.ListCount = 1 And f!cboRun = run, "Scoped route selection"
Check InStr(f!cboRun.Column(1), "Saved revision 1") > 0, "Readable revision label"
Check InStr(f!lstSteps.Column(1, 1), "Reviewer") > 0, "Reviewer name missing"
Check f!cmdAccept.Enabled And Not f!cmdRestart.Enabled, "Current step actions"
f!lstSteps = 1: Check UXReviewStep() And InStr(f!txtStep, "Read all saved blocks") > 0, "Full step instructions"
Call SwitchContributor(other): Call UXReviewStatus
Check Not f!cmdAccept.Enabled And Not f!cmdCancel.Enabled, "Contributor guards"
Call SwitchContributor(author)
Call CommitDraft(id, "EVAL", "Review fixture", "Synthetic member", #1/1/2026#, #9/30/2026#, "Changed", "Test", "")
Call UXReviewStatus
Check f!cmdRestart.Enabled And Not f!cmdAccept.Enabled, "Stale guidance"
Check InStr(f!lblStatus.Caption, "Needs restart") > 0, "Stale next action missing"
Call CommitDraft(id, "EVAL", "Review fixture", "Synthetic member", #1/1/2026#, #9/30/2026#, "Changed again", "Test", "")
Check Not UXReviewAction("Restart"), "Restart ignored changed display baseline"
Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending", "Failed restart modified old route"
Call UXReviewStatus
CurrentDb.Execute "UPDATE RouteTemplate SET IsArchived=True WHERE TemplateID=" & template
Check Not UXReviewAction("Restart"), "Archived template restarted"
Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending" And DCount("*", "RouteRun") = 1, "Restart rollback lost route"
CurrentDb.Execute "UPDATE RouteTemplate SET IsArchived=False WHERE TemplateID=" & template
CurrentDb.Execute "DELETE FROM RouteTemplateStep WHERE TemplateID=" & template
Call UXReviewStatus
Check Not UXReviewAction("Restart"), "Empty template restarted"
Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending" And DCount("*", "RouteRun") = 1, "Atomic replacement failure modified old route"
CurrentDb.Execute "INSERT INTO RouteTemplateStep (TemplateID,StepOrdinal,AssigneeID,Instructions) VALUES (" & template & ",1," & author & ",'Read all saved blocks')"
Call UXReviewStatus
If Not UXReviewAction("Restart") Then Err.Raise vbObjectError + 5950, , "Restart failed: " & f!lblStatus.Caption
Check True, "Restart succeeded"
newRun = f!cboRun
Check newRun <> run And DLookup("RunState", "RouteRun", "RunID=" & run) = "Cancelled", "Restart did not retain old history"
Check DLookup("RevisionID", "RouteRun", "RunID=" & newRun) = CurrentRevision(id) And RouteIsCurrent(newRun), "Restart revision wrong"
Check DLookup("StepState", "RouteRunStep", "RunID=" & newRun) = "Pending", "Restart reused decision"
DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
Forms!frmEditor!txtNotes = "Preserved pending notes"
f!txtNote = "Reviewed": Check UXReviewAction("Accept"), "Accept failed"
Check f!cmdApprove.Enabled And Not f!cmdAccept.Enabled, "Accepted next action"
f!txtNote = "Local only": Check UXReviewAction("Approve"), "Approval failed"
Check f!cmdReopen.Enabled And Not f!cmdStart.Enabled And Not f!cmdRestart.Enabled, "Approval lock"
Check Not Forms!frmEditor!cmdSave.Enabled And Forms!frmEditor!txtNotes = "Preserved pending notes", "Approval did not refresh editor lock or lost edits"
f!txtNote = "Correction": Check UXReviewAction("Reopen") And f!cmdStart.Enabled, "Reopen guidance"
Check CLng(f!txtReportID) = id And f!cboRun.ListCount = 2, "Actions lost report scope"
Check Forms!frmEditor!cmdSave.Enabled And EditorHasPending(), "Reopen did not restore editor access"
Check UXLeaveEditor("Discard"), "Pending review editor cleanup"
DoCmd.Close acForm, "frmUXReview", acSaveNo
id = CommitDraft(0, "EVAL", "Group review fixture", "Group member", Null, Null, "", "Test", "")
extraID = CommitDraft(0, "EVAL", "Group extra", "Extra member", Null, Null, "", "Test", "")
groupID = CreateSummaryGroup("Review group", "EVAL", "E5")
Call AddGroupMember(groupID, id, "E5")
run = StartRoute(id, template, groupID)
Call AddGroupMember(groupID, extraID, "E5")
DoCmd.OpenForm "frmUXReview", , , , , , CStr(id) & ":" & groupID
Set f = Forms!frmUXReview
Check f!cmdRestart.Enabled And Not RouteIsCurrent(run), "Group membership did not stale route"
Check UXReportState(id) = "Review needs restart", "Reports status missed stale group route"
Call SwitchContributor(other): Call UXReviewStatus
Check Not f!cmdRestart.Enabled And Not UXReviewAction("Restart"), "Wrong requester restarted"
Call SwitchContributor(author): Call UXReviewStatus
Check UXReviewAction("Restart"), "Group restart failed"
newRun = f!cboRun
Check DLookup("GroupID", "RouteRun", "RunID=" & newRun) = groupID And RouteIsCurrent(newRun), "Group context lost"
Call CommitDraft(id, "EVAL", "Group review fixture", "Group member", Null, Null, "Group stale", "Test", "")
CurrentDb.Execute "UPDATE SummaryGroup SET IsArchived=True WHERE GroupID=" & groupID
Call UXReviewStatus
Check Not UXReviewAction("Restart") And DLookup("RunState", "RouteRun", "RunID=" & newRun) = "Pending", "Archived group restart mutated route"
DoCmd.Close acForm, "frmUXReview", acSaveNo
For i = 1 To 90
Call CommitDraft(0, "EVAL", "Return fixture " & i, "Synthetic " & i, Null, Null, "", "Test", "")
Next
DoCmd.OpenForm "frmLibrary": Set f = Forms!frmLibrary
Dim previousFolder As Variant
previousFolder = f!lstFolders
f!uxFolderFilter = "Not a listed folder"
Check Not UXFolderFilterChanged(), "Unknown folder text accepted"
Check Nz(f!lstFolders, 0) = Nz(previousFolder, 0) And InStr(f!lblCount.Caption, "Choose a folder") > 0, "Unknown folder changed filter or lacked guidance"
DoCmd.OpenForm "frmEditor", , , , , , "0"
Forms!frmEditor!cboKind = "FITREP": Call UXEditorKindChanged
Forms!frmEditor!txtTitle = "Single result reopen fixture": Forms!frmEditor!txtOfficialName = "TESTER, UX"
Check UXSaveEditor(), "New FITREP save failed"
id = Forms!frmEditor!txtReportID
Check id > 0 And CurrentRevision(id) > 0, "New report did not create saved revision"
Check UXLeaveEditor(), "New report exit failed"
f!txtSearch = "Single result reopen fixture": Call RefreshLibrary
f!lstReports = id
Check UXReportScroll() = 0 And EditSelected(), "Single result at top cannot open"
Check Forms!frmEditor!txtOfficialName = "TESTER, UX", "New saved values not reopened"
Check UXLeaveEditor(), "Single result reopen exit failed"
f!txtSearch = "Return fixture": Call RefreshLibrary
f!lstReports = f!lstReports.Column(0, 50)
id = f!lstReports: position = UXReportScroll(30): topID = f!lstReports.Column(0, position)
Check position = 30, "Native scroll positioning failed"
Check EditSelected(), "Open selected failed"
Forms!frmEditor!txtNotes = "Auto save return"
Check UXLeaveEditor(), "Auto save failed"
deadline = Timer + 1: Do While Timer < deadline: DoEvents: Loop
Check Nz(f!lstReports, 0) = id And f!txtSearch = "Return fixture", "Selection/search lost"
Check CLng(f!lstReports.Column(0, UXReportScroll())) = topID, "Top visible report lost after reorder"
Check DLookup("DraftNotes", "Report", "ReportID=" & id) = "Auto save return", "Return did not save"
Check EditSelected(), "Second open failed"
CurrentDb.Execute "UPDATE Report SET IsArchived=True WHERE ReportID=" & id
Check UXLeaveEditor("Discard"), "Archived leave failed"
deadline = Timer + 1: Do While Timer < deadline: DoEvents: Loop
Check IsNull(f!lstReports), "Filtered-out report selected another report"
Check f!cboArchive = "Active", "Archive filter changed"
f!txtSearch = "No matching synthetic report": Call RefreshLibrary
Check f!lstReports.ListCount <= 1, "Empty filter failed"
Call UXCaptureReportList: Check UXRestoreReportList(), "Empty viewport restore failed"
TestReviewReturn = "PASS: " & checks & " focused review and native Reports return assertions": Exit Function
Failed:
TestReviewReturn = "FAIL after " & checks & ": " & Err.Description
End Function
