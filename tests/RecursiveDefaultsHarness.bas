Attribute VB_Name = "RecursiveDefaultsHarness"
Option Explicit
Private n As Long
Private Sub TraceStage(ByVal value As String)
 Dim h As Integer
 h = FreeFile: Open CurrentProject.Path & "\defaults-switch-trace.log" For Append As #h
 Print #h, value: Close #h
End Sub
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8210, , message
 n = n + 1
End Sub
Private Function Edit(ByVal key As String, ByVal value As Variant, ByVal decision As VbMsgBoxResult) As Boolean
 Forms!frmUXBulk.Controls("uxValue" & key) = value
 Forms!frmUXBulk!uxBlocks.Form.Controls("value" & key) = value
 Edit = UXConfirmBlockChange(key, decision)
End Function
Private Function TryApply(ByVal plan As Object, ByVal reopen As Boolean) As Boolean
 On Error GoTo Rejected
 ApplyRecursiveDefault plan, reopen: TryApply = True
Rejected:
End Function
Public Function TestRecursiveDefaults() As String
 On Error GoTo Failed
 Dim ga As Long, gb As Long, gc As Long, gone As Long, a As Long, b As Long, c As Long, orphan As Long, archived As Long, oldReport As Long
 Dim route As Long, template As Long, rev As Long, count As Long, events As Long, groups As Long, freshGroup As Long, freshReport As Long, stage As String
 Dim f As Form, key As Variant, name As Variant, db As DAO.Database, plan As Object, stamp As String
 ConfigureWorkspace "SYNTHETIC RECURSIVE DEFAULTS", "Tester"
 ga = CreateSummaryGroup("Alpha", "EVAL", "E5"): gb = CreateSummaryGroup("Bravo", "CHIEFEVAL", "E7"): gc = CreateSummaryGroup("Charlie", "FITREP", "O3")
 a = CommitDraft(0, "EVAL", "Alpha report", "A", Null, Null, "", "Test", "")
 b = CommitDraft(0, "CHIEFEVAL", "Bravo report", "B", Null, Null, "", "Test", "")
 c = CommitDraft(0, "FITREP", "Charlie report", "C", Null, Null, "", "Test", "")
 AddGroupMember ga, a, "E5": AddGroupMember gb, b, "E7": AddGroupMember gc, c, "O3"
 orphan = CommitDraft(0, "EVAL", "Ungrouped", "Orphan", Null, Null, "", "Test", "")
 archived = CommitDraft(0, "EVAL", "Archived member", "Archived", Null, Null, "", "Test", ""): AddGroupMember ga, archived, "E5"
 gone = CreateSummaryGroup("Archived group", "EVAL", "E5")
 oldReport = CommitDraft(0, "EVAL", "Archived group member", "Old", Null, Null, "", "Test", ""): AddGroupMember gone, oldReport, "E5"
 Set db = CurrentDb: db.Execute "UPDATE Report SET IsArchived=True WHERE ReportID=" & archived, dbFailOnError
 db.Execute "UPDATE SummaryGroup SET IsArchived=True WHERE GroupID=" & gone, dbFailOnError
 SetFieldDefault "Command", 1, "ShipStation", "OLD": SetFieldDefault "Group", ga, "ShipStation", "CUSTOM"
 template = CreateRouteTemplate("Synthetic route", CStr(MajorAuthor()), "Test")
 route = StartRoute(b, template, gb): DecideRouteStep route, "Accepted", "Test": ApproveSavedRoute route, "Test"
 stage = "Open and control cleanup": TraceStage stage
 Check UXOpenGlobalDefaults(), "Global defaults failed to open"
 Set f = Forms!frmUXBulk
 Check UXSharedScope(), "Scope banner failed before field selection"
 Check InStr(f!uxScope.Caption, "Scope: all ") = 1 And InStr(f!uxScope.Caption, "locally approved") > 0, "Scope banner missing counts or approval warning"
 'A malformed member list fails the advisory count; it must not raise or lose the cause.
 stage = "Scope failure": TraceStage stage
 'Access accepts a trailing comma in IN lists; an unknown name fails the query.
 Check UXBeginImmediateBlocks(0, False, "SyntheticBadMember"), "Scope failure raised into its caller"
 Check Not UXSharedScope(), "Malformed scope falsely succeeded"
 Set plan = SwiftLastFailureContext()
 Check plan("Procedure") = "UXSharedScope" And plan("Stage") = "Count scope reports" And plan("Number") <> 0, "Scope failure lost its primary error or stage"
 Check InStr(f!uxScope.Caption, "Scope unavailable.") = 1 And InStr(f!uxScope.ControlTipText, "UXSharedScope") > 0, "Scope failure diagnostic not visible"
 Check UXBeginImmediateBlocks(0, True, "") And InStr(f!uxScope.Caption, "Scope: all ") = 1, "Scope retry failed after restoring context"
 For Each name In Array("cmdPreview", "cmdApply", "lstPreview", "txtFullValues", "txtGroupImpact", "lblReason", "txtReason")
  Check Not UXHasControl(f, CStr(name)), "Retired preview control remains: " & name
 Next
 For Each key In UXSharedKeys()
  For Each name In Array("reset", "clear", "state")
   Check Not UXHasControl(f!uxBlocks.Form, CStr(name) & key), "Per-field legacy control remains"
  Next
 Next
 stage = "Cancel": TraceStage stage
 rev = CurrentRevision(a): count = DCount("*", "ReportRevision")
 Check Not Edit("ShipStation", "CANCELLED", vbNo), "Cancel reported applied"
 Check f!uxBlocks.Form!valueShipStation = "OLD", "Cancel did not restore UI"
 Check DCount("*", "ReportRevision") = count And UXGroupDefaultValue(0, "ShipStation") = "OLD", "Cancel wrote data"
 stage = "Global recursive apply": TraceStage stage
 Set plan = RecursiveDefaultPlan(0, True, "", "ShipStation", "GLOBAL")
 Check plan("Reports").Count = 3 And plan("Approvals") = 1, "Wrong recursive scope"
 Check InStr(RecursiveDefaultConfirmation(plan), "1 locally approved") > 0, "Approval reopening not disclosed"
 Check Not TryApply(plan, False), "Approval reopened without confirmation"
 Check Edit("ShipStation", "GLOBAL", vbYes), "Global change failed: " & f!uxStatus.Caption
 Check UXGroupDefaultValue(0, "ShipStation") = "GLOBAL", "Global default not saved"
 For Each key In Array(ga, gb, gc): Check UXGroupDefaultValue(CLng(key), "ShipStation") = "GLOBAL", "Child default not replaced": Next
 For Each key In Array(a, b, c): Check DLookup("ShipStation", "Report", "ReportID=" & key) = "GLOBAL", "Report not updated": Next
 Check CurrentRevision(a) <> rev And DLookup("WorkflowState", "Report", "ReportID=" & b) = "Draft", "Revision/approval state wrong"
 Check DLookup("RunState", "RouteRun", "RunID=" & route) = "Reopened", "Approval history not retained"
 For Each key In Array(orphan, archived, oldReport): Check IsNull(DLookup("ShipStation", "Report", "ReportID=" & key)), "Excluded report changed": Next
 Check DCount("*", "ChangeBatch") = 0 And Not UXGroupBulkPending(), "Preview stage was created"
 stage = "Failure rollback and retry": TraceStage stage
 route = StartRoute(b, template, gb): DecideRouteStep route, "Accepted", "Test": ApproveSavedRoute route, "Test"
 count = DCount("*", "ReportRevision"): events = DCount("*", "ReportEvent"): stamp = UXDefaultsStamp(ga)
 db.TableDefs("Report").ValidationRule = "[ReportID]<>" & b & " Or [ShipStation]<>'FAIL'"
 Check Not Edit("ShipStation", "FAIL", vbYes), "Injected failure succeeded"
 Check f!uxBlocks.Form!valueShipStation = "GLOBAL", "Failure did not restore field"
 Check DCount("*", "ReportRevision") = count And DCount("*", "ReportEvent") = events, "Failure left revisions/events"
 Check UXGroupDefaultValue(0, "ShipStation") = "GLOBAL" And UXDefaultsStamp(ga) = stamp, "Failure left defaults"
 Check DLookup("WorkflowState", "Report", "ReportID=" & b) = "Approved", "Failure lost approval"
 Check DLookup("ShipStation", "Report", "ReportID=" & a) = "GLOBAL", "Failure left earlier report change"
 db.TableDefs("Report").ValidationRule = ""
 Check Edit("ShipStation", "RETRY", vbYes), "Retry failed: " & f!uxStatus.Caption
 stage = "Invalid value and stale plan": TraceStage stage
 Check Not Edit("ShipStation", String$(19, "X"), vbYes), "Overlong value accepted"
 Check f!uxBlocks.Form!valueShipStation = "RETRY", "Invalid input stranded UI"
 Set plan = RecursiveDefaultPlan(0, True, "", "ShipStation", "STALE")
 SetFieldDefault "Group", ga, "ShipStation", "OTHER"
 Check Not TryApply(plan, True), "Stale plan accepted"
 Check UXGroupDefaultValue(0, "ShipStation") = "RETRY", "Stale failure changed global"
 UXCloseBlocks
 stage = "Future group inheritance": TraceStage stage
 freshGroup = CreateSummaryGroup("Future", "EVAL", "E5")
 freshReport = CreateGroupDraft(freshGroup, GroupStamp(freshGroup), UXDefaultsStamp(freshGroup), 0, 0, "Future member", "FUTURE, MEMBER", "IT2", "", Null, Null)
 Check DLookup("ShipStation", "Report", "ReportID=" & freshReport) = "RETRY", "Future report missed global default"
 stage = "Group-only apply": TraceStage stage
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = ga
 Check UXOpenGroupDefaults(), "Group defaults failed"
 Check Edit("ShipStation", "ALPHA", vbYes), "Group change failed"
 Check DLookup("ShipStation", "Report", "ReportID=" & a) = "ALPHA", "Group member unchanged"
 Check DLookup("ShipStation", "Report", "ReportID=" & b) = "RETRY" And UXGroupDefaultValue(0, "ShipStation") = "RETRY", "Group escaped scope"
 Check Edit("ComponentActive", True, vbYes), "Checkbox apply failed"
 Check DLookup("ComponentActive", "Report", "ReportID=" & a) = "Yes", "Checkbox not canonical"
 Check Edit("DateReported", #10/14/2026#, vbYes), "Date apply failed"
 Check IsDate(DLookup("DateReported", "Report", "ReportID=" & a)), "Date not saved"
 Check Edit("ShipStation", Null, vbYes), "Blank change failed"
 Check IsNull(DLookup("ShipStation", "Report", "ReportID=" & a)), "Blank did not clear report"
 Check Not UXGroupBulkPending(), "Change left staged state"
 UXCloseBlocks
 stage = "Empty group": TraceStage stage
 groups = CreateSummaryGroup("Empty", "EVAL", "E5"): Forms!frmMajorWorkbench!cboGroup = groups
 Check UXOpenGroupDefaults(), "Empty group defaults failed"
 Check Edit("ShipStation", "EMPTY", vbYes), "Empty group default failed"
 Check UXGroupDefaultValue(groups, "ShipStation") = "EMPTY", "Empty group default not saved"
 UXCloseBlocks
 stage = "Switch open defaults context without Close": TraceStage stage
 DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = ga
 Check UXListGroupDefaults(), "List group defaults failed"
 Check Forms!frmUXBulk!uxGroupContext = ga, "Initial group context wrong"
 Check SameFieldValue(Forms!frmUXBulk!uxValueShipStation, UXEffectiveDefault(ga, "ShipStation")), "Initial group values wrong"
 ' Returning to the list leaves the shared form open, as workspace tabs do.
 Forms!frmUXGroups.SetFocus
 Check UXOpenGlobalDefaults(), "Open group editor blocked global defaults"
 Check Forms!frmUXBulk!uxGroupContext = 0, "Global retained group context"
 Check Forms!frmUXBulk!lblHeading.Caption = "Global defaults", "Global heading stale"
 Check Not Forms!frmUXBulk!uxGroupTab1.Visible, "Global retained group tabs"
 Check Forms!frmUXBulk!uxValueShipStation = UXEffectiveDefault(0, "ShipStation"), "Global values stale"
 Check Edit("ShipStation", "SWITCH GLOBAL", vbYes), "Reopened global save failed"
 Check UXGroupDefaultValue(0, "ShipStation") = "SWITCH GLOBAL", "Reopened edit saved wrong scope"
 Forms!frmUXGroups.SetFocus: Forms!frmUXGroups!lstGroups = gb
 Check UXListGroupDefaults(), "Open global editor blocked selected group"
 Check Forms!frmUXBulk!uxGroupContext = gb And Forms!frmUXBulk!uxKind = "CHIEFEVAL", "Group context/type stale"
 Check Forms!frmUXBulk!uxGroupTab1.Visible, "Group tabs stayed hidden"
 Check Edit("ShipStation", "SWITCH BRAVO", vbYes), "Reopened group save failed"
 Check UXGroupDefaultValue(gb, "ShipStation") = "SWITCH BRAVO", "Reopened group default missing"
 Check UXGroupDefaultValue(0, "ShipStation") = "SWITCH GLOBAL", "Group edit escaped to global defaults"
 Check DLookup("ShipStation", "Report", "ReportID=" & b) = "SWITCH BRAVO", "Group switch missed member"
 Check DLookup("ShipStation", "Report", "ReportID=" & a) = "SWITCH GLOBAL", "Group switch changed other group"
 Forms!frmUXGroups.SetFocus: Forms!frmUXGroups!lstGroups = ga
 Check UXListGroupDefaults(), "Second group switch blocked"
 Check Forms!frmUXBulk!uxGroupContext = ga And Forms!frmUXBulk!uxKind = "EVAL", "Second group switch retained prior type"
 Check Forms!frmUXBulk!uxValueShipStation = "SWITCH GLOBAL", "Second group switch retained prior values"
 UXCloseBlocks
 TestRecursiveDefaults = "PASS: " & n & " direct defaults/control removal/cancel/scope/approval/rollback/retry/date/checkbox/future-report assertions"
 Exit Function
Failed:
 TestRecursiveDefaults = "FAIL: " & stage & " " & Err.Number & " " & Err.Description

End Function
