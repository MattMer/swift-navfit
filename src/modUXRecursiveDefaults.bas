Attribute VB_Name = "modUXRecursiveDefaults"
Option Explicit
Private confirmedValues As Object
Private directGroup As Long, directDefaults As Boolean, directCSV As String
' Capture the last committed control values so Cancel/failure restores the edit.
Public Function UXBeginImmediateBlocks(ByVal groupID As Long, ByVal defaults As Boolean, ByVal reportsCSV As String) As Boolean
 Dim key As Variant
 directGroup = groupID: directDefaults = defaults: directCSV = reportsCSV
 Set confirmedValues = CreateObject("Scripting.Dictionary")
 For Each key In UXSharedKeys(): confirmedValues(key) = Forms!frmUXBulk.Controls("uxValue" & key).Value: Next
 Call UXSharedScope()
 UXBeginImmediateBlocks = True
End Function
Private Function BlockCanonical(ByVal key As String, ByVal value As Variant) As Variant
 If UXFieldFlag(key) And Not IsNull(value) Then
  If key = "ObservationStatus" Then value = IIf(value, "NOB", "Observed") Else value = IIf(value, "Yes", "No")
 ElseIf UXFieldDate(key) Then
  value = UXDateEntry(value)
 End If
 BlockCanonical = value
End Function
' Build an in-memory plan; close every read cursor before any transaction writes.
Public Function RecursiveDefaultPlan(ByVal groupID As Long, ByVal defaults As Boolean, ByVal reportsCSV As String, ByVal key As String, ByVal value As Variant) As Object
 Dim plan As Object, groups As New Collection, reports As New Collection, values As Object, changes As Object, valid As Object, lifecycle As String
 Dim db As DAO.Database, r As DAO.Recordset, sql As String, ids As Object, item As Variant, csv As String, signature As String, approvals As Long, known As Boolean, signed As Long
 For Each item In UXSharedKeys(): If key = item Then known = True
 Next
 If Not known Then Err.Raise vbObjectError + 8200, "modUXRecursiveDefaults.RecursiveDefaultPlan", "This field is not a shared report block."
 If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 8201, "modUXRecursiveDefaults.RecursiveDefaultPlan", "Close the report editor before applying shared fields."
 Set db = CurrentDb
 If defaults And groupID = 0 Then
  Set r = db.OpenRecordset("SELECT GroupID FROM SummaryGroup WHERE IsArchived=False ORDER BY GroupID", dbOpenSnapshot)
  Do While Not r.EOF: groups.Add CLng(r!GroupID): r.MoveNext: Loop
  r.Close
 ElseIf groupID > 0 Then
  If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 8202, "modUXRecursiveDefaults.RecursiveDefaultPlan", "Recover the summary group before editing defaults."
  groups.Add groupID
 End If
 Set changes = CreateObject("Scripting.Dictionary"): changes(key) = value
 Set valid = UXBlockValidated(changes, "EVAL"): value = valid(key)
 signature = ""
 For Each item In groups: signature = signature & item & ":" & UXDefaultsStamp(CLng(item)) & "|": Next
 If defaults Then
  signature = signature & "ROOT:" & UXDefaultsStamp(groupID)
  sql = "SELECT DISTINCT r.ReportID,r.FormKind,r.CurrentRevisionID,r.WorkflowState FROM (Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID) INNER JOIN SummaryGroup AS g ON m.GroupID=g.GroupID WHERE r.IsArchived=False AND g.IsArchived=False"
  If groupID > 0 Then sql = sql & " AND m.GroupID=" & groupID
 Else
  Set ids = MajorIDs(reportsCSV)
  For Each item In ids.Keys
   If csv <> "" Then csv = csv & ","
   csv = csv & CLng(ids(item))
  Next
  If csv = "" Then Err.Raise vbObjectError + 8203, "modUXRecursiveDefaults.RecursiveDefaultPlan", "Select report members first."
  sql = "SELECT ReportID,FormKind,CurrentRevisionID,WorkflowState FROM Report WHERE IsArchived=False AND ReportID IN (" & csv & ")"
 End If
 Set r = db.OpenRecordset(sql & IIf(defaults, " ORDER BY r.ReportID", " ORDER BY ReportID"), dbOpenSnapshot)
 Do While Not r.EOF
  If UXFieldApplies(key, CStr(r!FormKind)) Then
   Set values = SavedFields(r!ReportID)
   signature = signature & r!ReportID & ":" & r!CurrentRevisionID & ":" & r!WorkflowState & "|"
   If Not SameFieldValue(values(key), value) Then
    'Signed reports are part of the binder: skip them and say so.
    lifecycle = ReportLifecycle(r!ReportID)
    If lifecycle = "Signed" Then
     signed = signed + 1
    Else
     If Not ReportTakesGroupChange(lifecycle) Then Err.Raise vbObjectError + 8204, "modUXRecursiveDefaults.RecursiveDefaultPlan", "Report " & r!ReportID & " is under review. Finish, pull back or cancel its review first; no changes were applied."
     reports.Add Array(CLng(r!ReportID), CLng(r!CurrentRevisionID), CStr(r!WorkflowState))
     If r!WorkflowState = "Approved" Then approvals = approvals + 1
    End If
   End If
  End If
  r.MoveNext
 Loop
 r.Close
 Set plan = CreateObject("Scripting.Dictionary")
 plan.Add "Groups", groups: plan.Add "Reports", reports: plan.Add "Approvals", approvals: plan.Add "SignedSkipped", signed
 plan.Add "GroupID", groupID: plan.Add "Defaults", defaults: plan.Add "CSV", reportsCSV
 plan.Add "Key", key: plan.Add "Value", value: plan.Add "Signature", signature
 Set RecursiveDefaultPlan = plan
End Function
Public Function RecursiveDefaultConfirmation(ByVal plan As Object) As String
 Dim scope As String, value As String
 If plan("Defaults") Then
  If plan("GroupID") = 0 Then scope = "all " & plan("Groups").Count & " active summary groups" Else scope = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & plan("GroupID")), "this summary group")
 Else
  scope = "the selected group members"
 End If
 value = Nz(plan("Value"), "(blank)"): If value = "" Then value = "(blank)"
 RecursiveDefaultConfirmation = "Apply " & UXFieldCaption(plan("Key"), Nz(Forms!frmUXBulk!uxKind, "EVAL")) & " = " & value & " recursively to " & scope & " and " & plan("Reports").Count & " existing reports?" & vbCrLf & vbCrLf & "Existing customized values for this field will be replaced. Archived items are excluded."
 If plan("Approvals") > 0 Then RecursiveDefaultConfirmation = RecursiveDefaultConfirmation & vbCrLf & plan("Approvals") & " locally approved reports will return to In-Work; previous approvals stay in history."
 If plan("SignedSkipped") > 0 Then RecursiveDefaultConfirmation = RecursiveDefaultConfirmation & vbCrLf & plan("SignedSkipped") & " signed reports are skipped; reopen one in the editor to correct it."
End Function
' Defaults, approval reopening and report revisions commit together or all roll back.
Public Sub ApplyRecursiveDefault(ByVal plan As Object, ByVal allowReopen As Boolean)
 Dim fresh As Object, w As DAO.Workspace, item As Variant, changes As Object, operation As Long, trans As Boolean, failure As Long, message As String
 On Error GoTo Failed
 If plan("Approvals") > 0 And Not allowReopen Then Err.Raise vbObjectError + 8205, "modUXRecursiveDefaults.ApplyRecursiveDefault", "Approval reopening requires confirmation."
 Set fresh = RecursiveDefaultPlan(plan("GroupID"), plan("Defaults"), plan("CSV"), plan("Key"), plan("Value"))
 If fresh("Signature") <> plan("Signature") Then Err.Raise vbObjectError + 8206, "modUXRecursiveDefaults.ApplyRecursiveDefault", "The scope changed during confirmation. Edit the field again."
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 operation = BeginOperation("Recursive shared field", "Field " & plan("Key") & "; scope " & plan("GroupID") & "; reports " & plan("Reports").Count)
 If plan("Defaults") Then
  If plan("GroupID") = 0 Then SetFieldDefault "Command", 1, plan("Key"), plan("Value")
  For Each item In plan("Groups"): SetFieldDefault "Group", CLng(item), plan("Key"), plan("Value"): Next
 End If
 Set changes = CreateObject("Scripting.Dictionary"): changes(plan("Key")) = plan("Value")
 For Each item In plan("Reports")
  If item(2) = "Approved" Then ReopenApprovedReport CLng(item(0)), "Confirmed recursive shared-field change: " & plan("Key")
  Call CommitSavedFields(CLng(item(0)), CLng(item(1)), changes, "Confirmed recursive shared-field change: " & plan("Key"))
 Next
 FinishOperation operation, "Completed", "Applied recursively to " & plan("Reports").Count & " reports; value=" & Nz(plan("Value"), "(blank)")
 w.CommitTrans: trans = False
 Call UXEditorContextChanged()
 Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXRecursiveDefaults", "ApplyRecursiveDefault", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
 failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
 If trans Then SwiftRollback w, swiftPrimary
 On Error GoTo 0
 SwiftRaise swiftPrimary
End Sub
' AfterUpdate is the sole commit entry: one confirmation, no preview/apply stage.
Public Function UXConfirmBlockChange(ByVal key As String, Optional ByVal decision As VbMsgBoxResult = 0) As Boolean
 Dim c As Access.Control, plan As Object, value As Variant, message As String
 On Error GoTo Failed
 Set c = Forms!frmUXBulk.Controls("uxValue" & key)
 If SameFieldValue(c.Value, confirmedValues(key)) Then UXConfirmBlockChange = True: Exit Function
 value = BlockCanonical(key, c.Value)
 Set plan = RecursiveDefaultPlan(directGroup, directDefaults, directCSV, key, value)
 If decision = 0 Then decision = MsgBox(RecursiveDefaultConfirmation(plan), vbYesNo + vbQuestion + vbDefaultButton2, "Apply field recursively?")
 If decision <> vbYes Then GoTo Restore
 DoCmd.Hourglass True
 ApplyRecursiveDefault plan, True
 confirmedValues(key) = c.Value
 Forms!frmUXBulk!uxStatus.Caption = "Saved changes to " & plan("Reports").Count & " existing reports; " & plan("Approvals") & " local approvals reopened." & IIf(plan("SignedSkipped") > 0, " " & plan("SignedSkipped") & " signed reports skipped.", "")
 UXConfirmBlockChange = True
Finished:
 Call UXSharedScope(key)
 DoCmd.Hourglass False
 Exit Function
Restore:
 Forms!frmUXBulk!uxStatus.Caption = "Not saved - cancelled. Previous field value restored; no report revisions changed."
 c.Value = confirmedValues(key)
 Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key) = c.Value
 GoTo Finished
Failed:
 Dim restoreFailure As Object
 Set restoreFailure = SwiftFailure("modUXRecursiveDefaults", "UXConfirmBlockChange", "Restore confirmed input", Err.Number, Err.Description, Err.Source, "See original operation outcome")
 message = CStr(restoreFailure("Description"))
 On Error Resume Next
 c.Value = confirmedValues(key): Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key) = c.Value
 If Err.Number <> 0 Then SwiftCleanupError restoreFailure, "restore pending value", Err.Number
 On Error GoTo 0
 Forms!frmUXBulk!uxStatus.Caption = "Not saved: " & SwiftFailureText(restoreFailure)
 DoCmd.Hourglass False
End Function
' Apply the report-shaped layout without any reset/clear/inheritance controls.
Public Function UXImmediateBlockMode() As Boolean
 Dim f As Access.Form
 Set f = Forms!frmUXBulk
 f!uxBlocks.Visible = True
 Call UXBlockViewport
 UXImmediateBlockMode = True
End Function
' Catalog actions can replace several related fields with one confirmation and
' one revision per report. The batch preview is an internal transaction service,
' not a screen, button or separate user requirement.
Public Function UXConfirmReportFields(ByVal groupID As Long, ByVal fields As Object, ByVal reason As String, Optional ByVal decision As VbMsgBoxResult = 0) As Boolean
 Dim db As DAO.Database, r As DAO.Recordset, targets As New Collection, values As Object, key As Variant, changed As Boolean, csv As String, approved As Long, lifecycle As String
 Dim item As Variant, w As DAO.Workspace, trans As Boolean, batch As Long, failure As Long, message As String, signed As Long
 On Error GoTo Failed
 If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 8202, "modUXRecursiveDefaults.UXConfirmReportFields", "Recover this group first."
 If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 8201, "modUXRecursiveDefaults.UXConfirmReportFields", "Close the report editor first."
 Set fields = UXBlockValidated(fields, Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & groupID), "EVAL"))
 Set db = CurrentDb: Set r = db.OpenRecordset("SELECT r.ReportID,r.CurrentRevisionID,r.WorkflowState FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE r.IsArchived=False AND m.GroupID=" & groupID & " ORDER BY r.ReportID", dbOpenSnapshot)
 Do While Not r.EOF
  Set values = SavedFields(r!ReportID): changed = False
  For Each key In fields.Keys
   If Not SameFieldValue(values(key), fields(key)) Then changed = True
  Next
  If changed Then lifecycle = ReportLifecycle(r!ReportID)
  If changed And lifecycle = "Signed" Then signed = signed + 1: changed = False
  If changed Then
   If Not ReportTakesGroupChange(lifecycle) Then Err.Raise vbObjectError + 8204, "modUXRecursiveDefaults.UXConfirmReportFields", "Report " & r!ReportID & " is under review. Finish, pull back or cancel its review first; no changes were applied."
   targets.Add Array(CLng(r!ReportID), CLng(r!CurrentRevisionID), CStr(r!WorkflowState))
   If r!WorkflowState = "Approved" Then approved = approved + 1
   If csv <> "" Then csv = csv & ","
   csv = csv & r!ReportID
  End If
  r.MoveNext
 Loop
 r.Close: Set r = Nothing
 If targets.Count = 0 Then UXConfirmReportFields = True: Exit Function
 message = reason & " for " & targets.Count & " reports? Existing values in these fields will be replaced."
 If approved > 0 Then message = message & vbCrLf & approved & " locally approved reports will return to In-Work; approvals remain in history."
 If signed > 0 Then message = message & vbCrLf & signed & " signed reports are skipped."
 If decision = 0 Then decision = MsgBox(message, vbYesNo + vbQuestion + vbDefaultButton2, "Apply to group members?")
 If decision <> vbYes Then Exit Function
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 For Each item In targets
  If CurrentRevision(CLng(item(0))) <> item(1) Then Err.Raise vbObjectError + 8206, "modUXRecursiveDefaults.UXConfirmReportFields", "Report changed during confirmation."
  If item(2) = "Approved" Then ReopenApprovedReport CLng(item(0)), reason
 Next
 batch = PreviewReportChanges(csv, fields, reason): ApplyFieldBatch batch
 w.CommitTrans: trans = False
 Call UXEditorContextChanged(): UXConfirmReportFields = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXRecursiveDefaults", "UXConfirmReportFields", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
 failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
 If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
 If trans Then SwiftRollback w, swiftPrimary
 On Error GoTo 0
 SwiftRaise swiftPrimary
End Function

' Scope is advisory before editing; the existing confirmation recomputes exact changes.
Public Function UXSharedScope(Optional ByVal key As String = "") As Boolean
 Dim db As DAO.Database, r As DAO.Recordset, sql As String, title As String, count As Long, approved As Long, locked As Long, message As String, groups As Long, applies As Boolean, signedCount As Long
 Dim stage As String, scopeFailure As Object
 On Error GoTo Failed
 If Not CurrentProject.AllForms("frmUXBulk").IsLoaded Then Exit Function
 stage = "Count scope reports"
 Set db = CurrentDb
 If directDefaults Then
  sql = "SELECT DISTINCT r.ReportID,r.FormKind,r.WorkflowState FROM (Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID) INNER JOIN SummaryGroup AS g ON m.GroupID=g.GroupID WHERE r.IsArchived=False AND g.IsArchived=False"
  If directGroup > 0 Then sql = sql & " AND m.GroupID=" & directGroup
  If directGroup = 0 Then
   groups = DCount("*", "SummaryGroup", "IsArchived=False"): title = "all " & groups & " active groups"
  Else
   title = "this group and its defaults"
  End If
 Else
  sql = "SELECT ReportID,FormKind,WorkflowState FROM Report WHERE IsArchived=False AND ReportID In (" & IIf(Len(directCSV) = 0, "0", directCSV) & ")"
  title = "selected members only"
 End If
 Set r = db.OpenRecordset(sql, dbOpenSnapshot)
 Do While Not r.EOF
  applies = True
  If Len(key) > 0 Then applies = UXFieldApplies(key, CStr(r!FormKind))
  If applies Then
   count = count + 1
   If r!WorkflowState = "Approved" Then approved = approved + 1
   If r!WorkflowState = "Signed" Then signedCount = signedCount + 1
   If ReportLifecycle(r!ReportID) = "In review" Then locked = locked + 1
  End If
  r.MoveNext
 Loop
 r.Close: Set r = Nothing
 stage = "Show scope"
 message = "Scope: " & title & "; " & count & " active reports, " & approved & " locally approved."
 message = message & " Custom values may be replaced; changed approvals reopen after confirmation. Archived items excluded."
 If locked > 0 Then message = message & " " & locked & " reports under review block changes until their review ends."
 If signedCount > 0 Then message = message & " " & signedCount & " signed reports are skipped."
 Forms!frmUXBulk!uxScope.Caption = message: Forms!frmUXBulk!uxScope.ControlTipText = Left$(message, 255)
 UXSharedScope = True: Exit Function
Failed:
 Set scopeFailure = SwiftFailure("modUXRecursiveDefaults", "UXSharedScope", stage, Err.Number, Err.Description, Err.Source)
 SwiftClose r, scopeFailure, "r"
 'Advisory only: never raise into the confirmation that called us after a commit.
 On Error Resume Next
 Forms!frmUXBulk!uxScope.Caption = "Scope unavailable. Exact affected reports and approvals are checked before confirmation."
 Forms!frmUXBulk!uxScope.ControlTipText = Left$(SwiftFailureText(scopeFailure), 255)
 If Err.Number <> 0 Then SwiftCleanupError scopeFailure, "show scope feedback", Err.Number
 On Error GoTo 0
End Function
