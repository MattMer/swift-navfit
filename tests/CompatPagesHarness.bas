Attribute VB_Name = "CompatPagesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8400, , message
 n = n + 1
End Sub
Public Function TestCompatPages() As String
 Dim kind As Variant, id As Long, f As Access.Form, s As Access.Form, row As Variant, key As String, page As Long, before As Long, renders As Long, x As Long, y As Long, width As Long, groupID As Long, values As Object, block As Long, entry As Long, previous As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC FIXED PAGES", "Tester"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  id = CommitDraft(0, CStr(kind), "Synthetic fixed page", "SYNTHETIC, MEMBER", #1/1/2026#, #9/30/2026#, "", administration:=Array("SYNTHETIC, MEMBER", "", "", "", ""))
  DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
  Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
  Check Not UXHasControl(f, "uxPageNotes") And f!uxPageReview.Caption = "Report tools", "Redundant Notes button remains or Report tools is missing"
  Check UXFixedEnabled(), "Fixed report pages unavailable"
  Check f!uxSection.Caption = "Front Page", "Front page failed: " & f!uxError.Caption
  For page = 0 To 1
   Check UXSelectSection(IIf(page = 0, 0, 5)), "Page could not open: " & f!uxError.Caption
   For Each row In ReportPageRows(CStr(kind), page)
    key = Split(CStr(row), "|")(0)
    If key <> "@guidance" Then Check s.Controls(UXFieldControl(key)).Visible, "Field missing on correct page: " & kind & " " & key
   Next
  Next
  For Each row In Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39")
   Check ReportPageForField(CStr(kind), CStr(row)) = 0, "Trait not on Front Page: " & row
  Next
  UXSelectSection 0
  If kind = "CHIEFEVAL" Then
   For block = 33 To 39
    previous = s.Controls("cboTrait" & block).TabIndex
    For entry = 1 To 3
     key = "Trait" & block & "Comment" & entry
     Check s.Controls("txt" & key).TabIndex > previous, "Chief comment tab order does not follow its trait: " & key
     previous = s.Controls("txt" & key).TabIndex
     Check s.Controls("txt" & key).Controls("label" & key).Caption = "Comment " & entry, "Chief comment lacks its attached label"
     Check s.Controls("txt" & key).Width = 6876 And s.Controls("txt" & key).Height = 680, "Chief comment is not the printed width with room for two edited lines"
     Check s.Controls("txt" & key).FontName = "Times New Roman" And s.Controls("txt" & key).FontSize = 12, "Chief comment font differs from the report"
     f!uxFieldSurface.SetFocus: s.Controls("txt" & key).SetFocus
     Check s.ActiveControl.Name = "txt" & key, "Chief comment cannot receive keyboard focus: " & s.ActiveControl.Name
    Next
    If block < 39 Then Check previous < s.Controls("cboTrait" & block + 1).TabIndex, "Chief comments follow a later trait in tab order"
   Next
   For block = 33 To 39
    Check s.Controls("uxChiefGuidance" & block).ControlType = acLabel And s.Controls("uxChiefGuidance" & block).OnClick = "", "Chief guidance is interactive"
   Next
   Check Not s!uxTraitGuidance.Visible, "Chief still displays the detached bottom guidance panel"
  End If
  s!txtOfficialName = "PENDING, MEMBER": UXSurfaceEdited "OfficialName"
  before = DCount("*", "ReportRevision")
  UXSelectSection 5: UXSelectSection 0
  Check s!txtOfficialName = "PENDING, MEMBER" And EditorHasPending(), "Page switch lost pending value"
  Check DCount("*", "ReportRevision") = before, "Page switch saved revision"
  x = s!txtOfficialName.Left: y = s!txtOfficialName.Top: width = s!txtOfficialName.Width: renders = UXFixedRenderCount()
  UXResizeEditor 16000, 10000: UXResizeEditor 28000, 14500: UXResizeEditor 12000, 8000
  Check s!txtOfficialName.Left = x And s!txtOfficialName.Top = y And s!txtOfficialName.Width = width, "Window change reflowed paper fields"
  Check UXFixedRenderCount() = renders, "Window change rebuilt fields"
  Check f!uxRail.Visible And f!uxNav0.Visible And f!uxNav5.Visible, "Main navigation rail unavailable"
  Check f!uxFieldSurface.Left > f!uxRail.Left + f!uxRail.Width, "Report overlaps the main navigation rail"
  Check f!cmdSave.Left > f!uxRail.Left + f!uxRail.Width And f!uxErrorList.Left > f!uxRail.Left + f!uxRail.Width, "Report actions/errors overlap navigation"
  Check UXFixedOpenField("Trait39") And s!cboTrait39.Visible And f!uxSection.Caption = "Front Page", "Validation could not navigate to front-page trait"
  Check UXSelectSection(7) And UXToolsForm()!txtConversation.Visible And UXToolsForm()!txtMessage.Visible, "Local notes unavailable"
  Check Not UXToolsForm()!lstDrafts.Visible And Not UXToolsForm()!lstMessages.Visible, "Retired conversation and draft lists still shown"
  Check UXSelectSection(8) And f!cmdRoute.Visible, "Local review unavailable"
  Check UXSelectSection(9) And f!lstRevisions.Visible, "History unavailable"
  Check UXLeaveEditor("Discard"), "Unable to close unchanged/explicitly discarded report"
 Next
 groupID = CreateSummaryGroup("Closeout group", "EVAL", "E5")
 id = CommitDraft(0, "EVAL", "Closeout test", "SYNTHETIC, MEMBER", #1/1/2026#, #9/30/2026#, "")
 AddGroupMember groupID, id, "E5"
 Check GroupCloseout(groupID, 1, 1, #9/30/2026#, #9/30/2026#) = "2026-09-30", "Closeout date wrong"
 Check GroupCloseout(groupID, 2, 2, #8/31/2026#, #9/30/2026#) = "Mixed", "Mixed end dates hidden"
 Check GroupCloseout(groupID, 2, 1, #9/30/2026#, #9/30/2026#) = "Incomplete", "Missing end date hidden"
 DoCmd.OpenForm "frmUXGroups": Check UXGroupList(), "Group list query failed"
 Check Forms!frmUXGroups!lstGroups.ColumnCount = 8, "Closeout column absent"
 Check UXSortList("frmUXGroups", "lstGroups", 7), "Closeout date heading sort failed"
 TestCompatPages = "PASS: " & n & " fixed report-page, pending buffer, local workflow and closeout assertions": Exit Function
Failed:
 TestCompatPages = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function

'Inspect native label associations, not proximity or a static caption heuristic.
Public Function PrepareEditorResetFeedback() As String
 Dim id As Long, group As Long, s As Access.Form
 On Error GoTo Failed
 id = CommitDraft(0, "CHIEFEVAL", "SYNTHETIC RESET", "SYNTHETIC CHIEF", #1/1/2026#, #9/30/2026#, "")
 group = CreateSummaryGroup("SYNTHETIC RESET GROUP", "CHIEFEVAL", "E7"): AddGroupMember group, id, "E7"
 Check UXShowGroupMembers(group, True), "Reset group did not open"
 Check UXGridOpenReport(), "Reset report did not open from Members"
 Set s = Forms!frmEditor!uxFieldSurface.Form
 s!txtTrait33Comment1 = "Pending authored Chief text": Check UXSurfaceEdited("Trait33Comment1"), "Comment edit failed"
 s!txtStart = "raw invalid date": Check UXSurfaceEdited("PeriodStart"), "Raw date edit failed"
 Check CheckpointEditor(), "Reset preparation checkpoint failed"
 PrepareEditorResetFeedback = "PASS: editor reset preparation": Exit Function
Failed:
 PrepareEditorResetFeedback = "FAIL: " & Err.Number & " " & Err.Description
End Function
Private Function RejectResetBaseline() As Boolean
 On Error GoTo Rejected
 GuardEditorBaseline
 Exit Function
Rejected:
 RejectResetBaseline = (Err.Number = vbObjectError + 900)
End Function
Public Function VerifyEditorResetFeedback() As String
 Dim id As Long, expected As Long, saved As Long, recovery As Long, changes As Object, s As Access.Form, group As Long
 On Error GoTo Failed
 n = 0: id = Forms!frmEditor!txtReportID: expected = Forms!frmEditor!uxRecoveryRevision: group = Forms!frmEditor!uxReturnGroup
 Set s = Forms!frmEditor!uxFieldSurface.Form
 Check EditorHasPending(), "Reset discarded the pending baseline"
 s!cboTrait34 = "5.0": Check UXSurfaceEdited("Trait34"), "Trait change after reset failed"
 Check s!uxChiefGuidance34.Caption = CatalogueTraitDescription("CHIEFEVAL", 34, "5.0"), "Per-trait guidance did not update"
 Check s!txtTrait33Comment1 = "Pending authored Chief text", "Trait change overwrote authored comments"
 Set changes = CreateObject("Scripting.Dictionary"): changes.Add "Title", "SYNTHETIC EXTERNAL UPDATE"
 CommitSavedFields id, CurrentRevision(id), changes, "Synthetic stale baseline check"
 saved = CurrentRevision(id)
 Check saved <> expected And RejectResetBaseline(), "Reset accepted a newer baseline"
 Check UXFeedbackCloseEditor(), "Back to Members trapped invalid raw edits"
 Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Checkpointed editor stayed open"
 Check CurrentProject.AllForms("frmUXComparison").IsLoaded And Forms!frmUXComparison!uxGroupContext = group, "Reset lost the return group"
 Check CurrentRevision(id) = saved And SavedFields(id)("Title") = "SYNTHETIC EXTERNAL UPDATE", "Navigation overwrote the newer saved revision"
 recovery = CLng(DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id))
 Check DLookup("BaseRevisionID", "RecoveryDraft", "RecoveryDraftID=" & recovery) = expected, "Checkpoint adopted a newer revision after reset"
 Check DLookup("StartEntry", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "raw invalid date", "Navigation lost the invalid raw date"
 Check DLookup("Trait33Comment1", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "Pending authored Chief text", "Navigation lost the authored comment"
 Check RestoreRecoveryDraft(recovery, True), "Reset checkpoint could not be recovered as a new draft"
 Check Forms!frmEditor!txtTrait33Comment1 = "Pending authored Chief text" And Forms!frmEditor!txtStart = "raw invalid date", "Recovered raw values differ"
 Check UXLeaveEditor("Discard"), "Explicit synthetic recovery discard failed"
 VerifyEditorResetFeedback = "PASS: " & n & " real-reset trait/navigation/recovery and stale-baseline assertions": Exit Function
Failed:
 VerifyEditorResetFeedback = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function TestEntryLabels() As String
 Dim spec As Variant, f As Access.Form, c As Access.Control, checked As Long
 On Error GoTo Failed
 For Each spec In Array( _
  Array("frmSetup", "txtCommand", "lblCommand"), Array("frmSetup", "txtProfile", "lblUser"), _
  Array("frmContributorStart", "cboContributor", "lblExisting"), Array("frmContributorStart", "txtNewContributor", "lblNew"), _
  Array("frmPDFOutput", "cboSource", "lblSource"), Array("frmPDFOutput", "cboFormat", "lblFormat"), _
  Array("frmPDFOutput", "chkSummary", "lblSummary"), Array("frmPDFOutput", "txtFolder", "lblFolder"), _
  Array("frmPDFOutput", "lstReports", "lblReports"), Array("frmPDFOutput", "txtFindings", "lblFindings"))
  DoCmd.OpenForm CStr(spec(0)), acDesign
  Set f = Forms(CStr(spec(0))): Set c = f.Controls(CStr(spec(1)))
  Check Len(c.Controls(CStr(spec(2))).Caption) > 0, "Unlabelled entry control: " & spec(0) & "." & spec(1)
  Check c.TabStop, "Entry control is not in the tab sequence: " & spec(1)
  checked = checked + 1
  DoCmd.Close acForm, CStr(spec(0)), acSaveNo
 Next
 TestEntryLabels = "PASS: " & checked & " native entry-control label associations and tab stops": Exit Function
Failed:
 TestEntryLabels = "FAIL: entry labels: " & Err.Number & " " & Err.Description
End Function
