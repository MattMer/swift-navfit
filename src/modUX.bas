Attribute VB_Name = "modUX"
Option Compare Binary
Option Explicit
Private editorGroupRefreshes As Long
Private editorContextDirty As Boolean
Private editorContextReport As Long
Private editorContextGroup As Long
Private currentSection As Long
Private currentPage As Long
Private editorStarting As Boolean
Private bulkChanges As Object
Private bulkCSV As String
Private bulkBatch As Long
Private groupContext As Long
Private editorCloseAuthorized As Boolean
Private editorCloseBusy As Boolean
Private bulkPreviewReason As String
Private bulkPreviewStamp As String
Private bulkTargetStamp As String
Private bulkDefaultBefore As Object
Private bulkDefaultPreview As Object
Private bulkMixed As Object
Private bulkDefaults As Boolean
Private bulkDefaultStamp As String
Private bulkDefaultResets As Object

'Required live-view contracts are distinct from optional compatibility probes.
Public Const UX_EDITOR_VIEW As String = "frmEditor"
Public Const UX_MEMBERS_VIEW As String = "frmUXComparison"
Public Const UX_TRAIT_ROWS_VIEW As String = "frmUXTraitRows"
' One quoted value-list cell. Access splits an unquoted cell at commas as well
' as semicolons, so quoting keeps names and titles in their own column.
Public Function UXListCell(ByVal value As Variant) As String
    UXListCell = """" & Replace(Replace(Replace(Nz(value, ""), """", """"""), vbCr, " "), vbLf, " ") & """"
End Function
Public Function UXViewContracts(ByVal view As String) As Variant
    Select Case view
        Case "frmUXCounselingDate": UXViewContracts = Array("txtDate", "cmdSave|OnClick|=UXCounselingCalendarUse()", "cmdCancel", "lblStatus")
        Case UX_EDITOR_VIEW: UXViewContracts = Array("uxPreviousReport", "uxNextReport", "cmdSave|OnClick|=UXSaveEditor()", "txtReportID", "cboKind", "lblSaved", "lblRecovery", "uxFrontPage", "uxBackPage", "uxToolsSurface", "uxRecoverySession", "uxRecoveryRevision", "uxReturnGroup")
        Case UX_MEMBERS_VIEW: UXViewContracts = Array("uxTraitGrid|SourceObject|Form.frmUXTraitRows", "uxGroupContext", "txtMeeting", "lblStatus", "cmdRetryCells|OnClick|=UXGridDrain()", "cmdDiscardCells|OnClick|=UXGridDiscardPending()", "uxCurrentExport|OnClick|=UXCurrentGroupTool(""frmUXComparison"",""Export"")", "uxCurrentPDF|OnClick|=PDFOutputOpenGroup(""frmUXComparison"")")
        Case "frmMyWork": UXViewContracts = Array("lstWork|AfterUpdate|=UXWorkSelection()", "lstWork|OnDblClick|=UXWorkOpen()", "cmdOpenWork|OnClick|=UXWorkOpen()", "cmdDraft|OnClick|=UXWorkOpen(True)", "cmdRefresh|OnClick|=LoadMyWork()", "uxContributor", "chkAll", "txtHistory", "txtNotes", "uxWorkStatus", "cmdAccept|OnClick|=FinishReviewUI(""Accepted"")", "cmdReturn|OnClick|=FinishReviewUI(""Returned"")", "cmdCancel|OnClick|=FinishReviewUI(""Cancelled"")")
        Case "frmUXNativeGroups": UXViewContracts = Array("cmdPreview|OnClick|=UXPreviewNativeGroups()", "cmdApply|OnClick|=UXApplyNativeGroups()", "cmdBack|OnClick|=UXCloseNativeGroups()", "cmdAssignImportGrade|OnClick|=UXAssignNativeGroupGrade()", "lstGroups|OnClick|=UXNativeGroupSelectionChanged()", "txtSource", "txtPreview", "lblStatus", "cboImportGrade")
        Case "frmUXMoveGroup": UXViewContracts = Array("cmdPreview|OnClick|=UXPreviewMoveGroup()", "cmdApply|OnClick|=UXApplyMoveGroup()", "cmdClose|OnClick|=UXCloseScreen(""frmUXMoveGroup"")", "reportID", "cboGroup", "txtPreview", "lblStatus", "lblMember")
        Case "frmUXDirectoryMerge": UXViewContracts = Array("cmdPreview|OnClick|=UXDirectoryMergePreview()", "cmdApply|OnClick|=UXDirectoryMergeApply()", "cmdUndo|OnClick|=UXDirectoryMergeUndo()", "cmdClose|OnClick|=UXDirectoryClose(""frmUXDirectoryMerge"")", "cboSource|AfterUpdate|=UXDirectoryMergeChanged()", "cboTarget|AfterUpdate|=UXDirectoryMergeChanged()", "lstCandidates|OnDblClick|=UXDirectoryMergeCandidate()", "txtPreview", "txtReason", "lstHistory")
        Case UX_TRAIT_ROWS_VIEW: UXViewContracts = Array("ReportID", "RevisionID", "Member", "RowStatus", "PromotionRecommendation", "AssignmentRecommendationA", "AssignmentRecommendationB")
        Case "frmPDFImport": UXViewContracts = Array("cmdResume|OnClick|=UXResumePDFImport()", "cmdCancelPreview|OnClick|=UXCancelPDFImport()", "cmdImport|OnClick|=ImportPDFUI()", "cmdPreview|OnClick|=UXPreviewPDFImport()", "txtPreview", "cboIdentifier", "cboTarget", "cboMergeMode", "chkLegacy", "cmdOpenResult", "uxState", "lblStatus")
        Case "frmNAVFITExchange": UXViewContracts = Array("cmdExportGroup|OnClick|=UXExportNAVFITGroup()", "cmdtxtExportPath|OnClick|=UXChooseNAVFITGroupFolder()", "cmdBack|OnClick|=UXCloseScreen(""frmNAVFITExchange"")", "cboGroup", "lblGroup", "txtExportPath", "txtStatus", "txtTemplate")
        Case "frmUXWorkspaceImport": UXViewContracts = Array("cmdResume|OnClick|=UXResumeSwiftWorkspace()", "cmdCancelPreview|OnClick|=UXCancelSwiftWorkspacePreview()", "cmdApply|OnClick|=UXApplySwiftWorkspace()", "cmdClose|OnClick|=UXCloseScreen(""frmUXWorkspaceImport"")", "txtPreview", "txtProgress", "lblStatus")
        Case "frmPDFOutput": UXViewContracts = Array("cmdGenerate|OnClick|=PDFOutputCreate()", "cmdPagePreview|OnClick|=PDFOutputPagePreview()", "cmdBrowse|OnClick|=PDFOutputBrowseFolder()", "cmdRefresh|OnClick|=PDFOutputRefreshPreview()", "cmdClose|OnClick|=PDFOutputClose()", "cmdOpenPDF|OnClick|=PDFOutputOpenResult(False)", "cmdOpenFolder|OnClick|=PDFOutputOpenResult(True)", "lstReports|AfterUpdate|=PDFOutputSelectRow()", "txtFolder", "txtFindings", "cboSource", "cboFormat", "lblStatus")
        Case "frmLibrary": UXViewContracts = Array("txtPeriodFrom|AfterUpdate|=RefreshLibrary()", "txtPeriodThrough|AfterUpdate|=RefreshLibrary()", "cboSavedFilter|AfterUpdate|=UXLoadReportFilter()", "cmdSaveFilter|OnClick|=UXSaveReportFilter()", "cmdDeleteFilter|OnClick|=UXDeleteReportFilter()")
        Case "frmUXMemberHistory": UXViewContracts = Array("chkArchivedCounseling|AfterUpdate|=UXHistoryCounselingLoad()", "cmdCounselingArchive|OnClick|=UXCounselingArchive()")
        Case Else: UXViewContracts = Array()
    End Select
End Function
'Finite dynamic families, derived from the same field/key catalogues as their
'callers. The file-only build checker expands these families independently.
Public Function UXDynamicViewControls(ByVal view As String) As Collection
    Dim result As New Collection, keys As Variant, key As Variant, i As Long, routes As Variant
    keys = Array()
    Select Case view
        Case "frmMemberDetails": keys = MemberDetailsKeys()
        Case "frmDecisionDetails": keys = DecisionDetailsKeys()
        Case "frmPartyNames": keys = PartyNamesKeys()
        Case "frmReportingSenior": keys = ReportingSeniorKeys()
        Case "frmReportPeople": keys = ReportPeopleKeys()
        Case "frmServiceDetails": keys = ServiceDetailsKeys()
        Case "frmContextOccasions", "frmContextCategories", "frmContextDuties": keys = ReportContextPanelKeys(Mid$(view, 11))
    End Select
    For Each key In keys: result.Add "txt" & key: Next
    If view = "frmUXSeniorEditor" Then
        For Each key In RSRecordKeys(): result.Add "rs" & key: Next
    End If
    If view = "frmUXRSCATable" Then
        For Each key In RSTableGrades()
            result.Add "grade" & key: result.Add "avg" & key: result.Add "count" & key: result.Add "date" & key: result.Add "src" & key
        Next
    End If
    If view = "frmUXBulk" Or view = "frmUXSharedFields" Then
        For Each key In UXSharedKeys()
            If view = "frmUXBulk" Then result.Add "uxValue" & key Else result.Add "value" & key: result.Add "label" & key
        Next
    End If
    If view = UX_EDITOR_VIEW Then
        For i = 0 To 9: result.Add "uxSection" & i: Next
    End If
    If view = UX_EDITOR_VIEW Or view = "frmUXPriorEditor" Then
        For i = 33 To 39: result.Add "lblTrait" & i: result.Add "cboTrait" & i: Next
    End If
    If view = UX_TRAIT_ROWS_VIEW Then
        For i = 33 To 39: result.Add "heading" & i: result.Add "Trait" & i: Next
    End If
    If view = "frmUXMemberEditor" Then
        keys = RUADKeys()
        For i = 0 To UBound(keys)
            If keys(i) <> "NAME" Then result.Add "ruad" & i: result.Add "ruadLabel" & i
        Next
        For i = 0 To 4: result.Add "ruadPage" & i: Next
    End If
    If view = "frmUXEditorFields" Then
        For Each key In CatalogueFieldKeys()
            If key <> "FormKind" Then result.Add UXFieldControl(CStr(key)): result.Add "label" & key: result.Add "err" & key
        Next
        For i = 0 To 2: result.Add "uxGroupBox" & i: Next
        For i = 33 To 39: result.Add "uxChiefGuidance" & i: Next
    End If
    If view <> UX_TRAIT_ROWS_VIEW And view <> "frmUXEditorFields" And view <> "frmUXSharedFields" And view <> "frmUXRSCATable" Then
        If UXHasControl(Forms(view), "uxCompactNav") Then
            routes = UXWorkspaceRoutes()
            For i = 0 To UBound(routes): result.Add "uxNav" & i: Next
        End If
        If UXHasControl(Forms(view), "uxGroupTab0") Then
            For Each key In Array(0, 1, 2, 4): result.Add "uxGroupTab" & key: Next
        End If
        Dim list As Access.Control
        For Each list In Forms(view).Controls
            If list.ControlType = acListBox Then
                If UXHasControl(Forms(view), "sort_" & list.Name & "_0") Then
                    For i = 0 To list.ColumnCount - 1: result.Add "sort_" & list.Name & "_" & i: Next
                End If
            End If
        Next
    End If
    Set UXDynamicViewControls = result
End Function
Public Function UXRequireControl(ByVal view As String, ByVal name As String) As Access.control
    On Error GoTo Missing
    If view = "frmUXSharedFields" And CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        If Forms!frmUXBulk.CurrentView <> 0 Then
            Set UXRequireControl = Forms!frmUXBulk!uxBlocks.Form.Controls(name): Exit Function
        End If
    End If
    If view = "frmUXRSCATable" And CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then
        If Forms!frmUXSeniorEditor.CurrentView <> 0 Then
            Set UXRequireControl = Forms!frmUXSeniorEditor!uxRSCATable.Form.Controls(name): Exit Function
        End If
    End If
    If view = "frmUXEditorFields" And CurrentProject.AllForms(UX_EDITOR_VIEW).IsLoaded Then
        If Forms(UX_EDITOR_VIEW).CurrentView <> 0 Then
            Set UXRequireControl = Forms(UX_EDITOR_VIEW)!uxFieldSurface.Form.Controls(name)
            Exit Function
        End If
    End If
    If view = UX_TRAIT_ROWS_VIEW And CurrentProject.AllForms(UX_MEMBERS_VIEW).IsLoaded Then
        If Forms(UX_MEMBERS_VIEW).CurrentView <> 0 Then
            Set UXRequireControl = Forms(UX_MEMBERS_VIEW)!uxTraitGrid.Form.Controls(name)
            Exit Function
        End If
    End If
    Set UXRequireControl = Forms(view).Controls(name)
    Exit Function
Missing:
    Err.Raise vbObjectError + 9400, "modUX.UXRequireControl", "Required view contract missing: " & view & "." & name & ". Restore a verified application copy."
End Function
Public Function UXAssertView(ByVal view As String) As Boolean
    Dim entry As Variant, parts As Variant, c As Access.Control, key As Variant, block As Long, actual As String, expected As String
    For Each entry In UXViewContracts(view)
        parts = Split(CStr(entry), "|"): Set c = UXRequireControl(view, CStr(parts(0)))
        If UBound(parts) = 2 Then
            actual = Nz(c.Properties(parts(1)).Value, ""): expected = CStr(parts(2))
            If parts(1) = "SourceObject" Then
                If Left$(actual, 5) = "Form." Then actual = Mid$(actual, 6)
                If Left$(expected, 5) = "Form." Then expected = Mid$(expected, 6)
            End If
            If StrComp(actual, expected, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9401, "modUX.UXAssertView", "Required view binding differs: " & view & "." & parts(0) & "." & parts(1)
        End If
    Next
    For Each key In UXDynamicViewControls(view): Set c = UXRequireControl(view, CStr(key)): Next
    If view = UX_EDITOR_VIEW Then
        For Each key In CatalogueControls(): Set c = UXRequireControl(view, CStr(key)): Next
    End If
    If view = UX_TRAIT_ROWS_VIEW Then
        For block = 33 To 39
            Set c = UXRequireControl(view, "Trait" & block)
            If c.AfterUpdate <> "=UXGridQueue(""Trait" & block & """)" Then Err.Raise vbObjectError + 9401, "modUX.UXAssertView", "Required trait queue binding differs: " & view & ".Trait" & block
        Next
    End If
    UXAssertView = True
End Function
'Return caught failures to automation rather than opening an unhandled VBA dialog.
Public Function UXCheckView(ByVal view As String) As String
    On Error GoTo Failed
    Call UXAssertView(view)
    UXCheckView = "PASS": Exit Function
Failed:
    UXCheckView = "FAIL: " & Err.Number & " " & Err.Description
End Function

'One workspace destination list for native dispatch and generated navigation rails.
Public Function UXWorkspaceRoutes() As Variant
    UXWorkspaceRoutes = Array("Overview|frmHome", "Reports|frmLibrary", "Summary groups|frmUXGroups", "My Work|frmMyWork", "Resources|frmUXResources", "Settings & Help|frmUXSettings")
End Function
Public Function UXWorkspaceRouteList() As String
    UXWorkspaceRouteList = Join(UXWorkspaceRoutes(), ";")
End Function
Public Function UXWorkspaceRowSource() As String
    Dim row As Variant, parts As Variant
    For Each row In UXWorkspaceRoutes()
        parts = Split(CStr(row), "|")
        If Len(UXWorkspaceRowSource) > 0 Then UXWorkspaceRowSource = UXWorkspaceRowSource & ";"
        UXWorkspaceRowSource = UXWorkspaceRowSource & parts(1) & ";" & parts(0)
    Next
End Function
Public Function UXIsWorkspace(ByVal target As String) As Boolean
    Dim row As Variant
    For Each row In UXWorkspaceRoutes()
        If Split(CStr(row), "|")(1) = target Then UXIsWorkspace = True: Exit Function
    Next
End Function
Public Function UXLegacyTool(ByVal target As String) As String
    Select Case target
        Case "frmReviewComments", "frmUXNoteComposer": UXLegacyTool = "Notes"
        Case "frmUXReview": UXLegacyTool = "Review"
        Case "frmChecklist": UXLegacyTool = "Validation"
    End Select
End Function
Public Function UXLegacyToolOpen(ByVal view As String) As Boolean
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        UXLegacyToolOpen = UXToolsChoose(view)
    Else
        UXLegacyToolOpen = UXNavigate("frmLibrary")
    End If
End Function
' Invalidate cached membership and saved group-average data after a local mutation;
' read it on the next editor refresh rather than polling on every field edit.
Public Function UXEditorContextChanged() As Boolean
    editorContextDirty = True: UXEditorContextChanged = True
End Function
' Expose context-refresh frequency for performance diagnostics without database reads.
Public Function UXEditorGroupRefreshCount() As Long
    UXEditorGroupRefreshCount = editorGroupRefreshes
End Function
' Apply the summary-group report filter. The control name is retained for legacy form compatibility.
Public Function UXFolderFilterChanged() As Boolean
    If Not IsNumeric(Nz(Forms!frmLibrary!uxFolderFilter, "")) Then
        Forms!frmLibrary!lblCount.Caption = "Choose a summary group from the list."
        Forms!frmLibrary!uxFolderFilter = Forms!frmLibrary!lstFolders
        Exit Function
    End If
    Forms!frmLibrary!lstFolders = Forms!frmLibrary!uxFolderFilter
    UXFolderFilterChanged = RefreshLibrary()
End Function
' Probe optional generated controls so adapters can tolerate form layout differences.
Public Function UXHasControl(ByVal f As Access.Form, ByVal name As String) As Boolean
    On Error GoTo Missing
    Dim c As Access.Control: Set c = f.Controls(name): UXHasControl = True
Missing:
End Function
' Store the modal decision and close the dialog for its waiting caller.
Public Function UXDecide(ByVal decision As String) As Boolean
    TempVars!UXDecision = decision
    DoCmd.Close acForm, "frmUXDecision", acSaveNo
    UXDecide = True
End Function
' Open a modal decision dialog, defaulting to Cancel if no choice is completed.
Public Function UXAsk(ByVal mode As String) As String
    On Error Resume Next: TempVars.Remove "UXDecision": On Error GoTo 0
    TempVars.Add "UXDecision", "Cancel"
    DoCmd.OpenForm "frmUXDecision", , , , , acDialog, mode
    UXAsk = CStr(TempVars!UXDecision): TempVars.Remove "UXDecision"
End Function
' Adapt decision wording for review, reporting-senior editing or report exit.
Public Function UXLoadDecision() As Boolean
    With Forms!frmUXDecision
        If Nz(.OpenArgs, "") = "Review" Then
            !lblHeading.Caption = "Use which report values?"
            !lblBody.Caption = "Review actions use saved revisions. Save pending changes first, use the saved version while keeping pending edits, or cancel."
            !cmdDiscard.Caption = "Use saved version"
        ElseIf Nz(.OpenArgs, "") = "Senior" Then
            !lblHeading.Caption = "Keep reporting senior changes?"
            !lblBody.Caption = "Save pending record details and any completed RSCA baseline, discard changes, or keep editing."
            !cmdDiscard.Caption = "Discard changes"
        Else
            !lblHeading.Caption = "Keep your report changes?"
            !lblBody.Caption = "Saving creates a report revision. Recovery checkpoints are separate. Discard removes this editor session's pending checkpoint."
            !cmdDiscard.Caption = "Discard changes"
        End If
    End With
    UXLoadDecision = True
End Function
' Flush visible pending values and save by default before leaving. A failed save
' returns False and preserves the buffer; explicit discard clears the session checkpoint.
Public Function UXPrepareEditorClose(Optional ByVal decision As String = "") As Boolean
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then UXPrepareEditorClose = True: Exit Function
    If editorCloseBusy Then Exit Function
    On Error GoTo Failed
    editorCloseBusy = True
    If Not UXToolsCheckpoint(True) Then GoTo Finished
    UXFlushSurface
    If EditorHasPending() Then
        If Len(decision) = 0 Then decision = "Save"
        Select Case decision
            Case "Save"
                If Not UXSaveEditor() Then
                    'Invalid raw values cannot enter typed saved fields, but must not
                    'trap navigation. Retain them in contributor-owned recovery.
                    If Not CheckpointEditor() Then GoTo Finished
                    UXPrepareEditorClose = True: GoTo Finished
                End If
            Case "Discard"
            Case Else: GoTo Finished
        End Select
    End If
    DiscardEditorCheckpoint
    UXPrepareEditorClose = True
Finished:
    editorCloseBusy = False
    Exit Function
Failed:
    UXShowSaveError "Unable to leave this report: " & Err.Description
    Resume Finished
End Function
' Authorize native form closure only after exit handling succeeds; checkpoint raw
' values when a failed save blocks normal closure.
Public Function UXEditorUnload() As Boolean
    If editorCloseAuthorized Then UXEditorUnload = True Else UXEditorUnload = UXPrepareEditorClose()
    ' Forced Access shutdown can ignore Cancel. Retain raw edits when saving fails.
    If Not UXEditorUnload Then CheckpointEditor Else PDFFieldFitRelease
    'The editor's print preview page cache goes with it (Create PDFs keeps its own).
    If UXEditorUnload And Not PDFOutputBusy() Then
        If Not CurrentProject.AllForms("frmPDFOutput").IsLoaded Then PDFOutputEndPreview
    End If
End Function
' Close the editor once exit handling succeeds, avoiding a duplicate unload save.
Public Function UXLeaveEditor(Optional ByVal decision As String = "") As Boolean
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then UXLeaveEditor = True: Exit Function
    If Not UXPrepareEditorClose(decision) Then Exit Function
    On Error GoTo Finished
    editorCloseAuthorized = True
    If CurrentProject.AllForms("frmUXCounselingDate").IsLoaded Then DoCmd.Close acForm, "frmUXCounselingDate", acSaveNo
    If CurrentProject.AllForms("frmUXValidation").IsLoaded Then DoCmd.Close acForm, "frmUXValidation", acSaveNo
    DoCmd.Close acForm, "frmEditor", acSaveNo
    UXLeaveEditor = True
Finished:
    editorCloseAuthorized = False
End Function
' Checkpoint and close the editor for backup without committing pending values as a revision.
Public Function UXCloseEditorForBackup() As Boolean
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then UXCloseEditorForBackup = True: Exit Function
    On Error GoTo Failed
    If Not CheckpointEditor() Then Exit Function
    editorCloseAuthorized = True
    DoCmd.Close acForm, "frmEditor", acSaveNo
    UXCloseEditorForBackup = True
Finished:
    editorCloseAuthorized = False
    Exit Function
Failed:
    UXShowSaveError "Unable to checkpoint and close for backup: " & Err.Description
    Resume Finished
End Function
' Leave contributor/report editors safely before opening the requested workspace.
Public Function UXNavigate(ByVal target As String) As Boolean
    On Error GoTo Failed
    If target = "frmBulkFolders" Then target = "frmLibrary"
    If Not HasActiveContributor() Then
        UXNavigate = ShowContributorStartup(): Exit Function
    End If
    If UXLegacyTool(target) <> "" Then UXNavigate = UXLegacyToolOpen(UXLegacyTool(target)): Exit Function
    If target = "frmHelp" Or target = "frmUXHelp" Then
        DoCmd.OpenForm "frmUXHelp"
        UXNavigate = True: Exit Function
    End If
    If Not RSLeaveSeniorEditor() Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    If Not UXLeaveGroup() Then Exit Function
    DoCmd.OpenForm target
    Forms(target).SetFocus
    UXLoadContext
    UXNavigate = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Unable to open this workspace"
End Function
' Initialize the report workspace at its first official section.
Public Function UXEditorStart() As Boolean
    On Error GoTo Finished
    editorStarting = True
    currentSection = 0: currentPage = 0
    UXEditorStart = UXSelectSection(0)
Finished:
    editorStarting = False
End Function
' Show the requested official/local section using the same unbound buffer. Section
' changes preserve pending values and do not create report revisions.
Public Function UXSelectSection(ByVal section As Long, Optional ByVal page As Long = 0) As Boolean
    Dim f As Access.Form, c As Access.Control, tags As Variant, id As Long, locked As Boolean, maxPage As Long
    If section = 1 Then section = 0
    Set f = Forms!frmEditor
    If UXFixedEnabled() Then
        currentSection = section: currentPage = page
        UXSelectSection = UXFixedSelect(section): Exit Function
    End If
    If UXHasControl(f, "cmdClose") Then f!cmdClose.SetFocus
    currentSection = section: currentPage = page: id = Nz(f!txtReportID, 0)
    locked = Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & id), False) Or Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & id), "Draft") <> "Draft"
    If UXSurfaceAvailable() Then
        For Each c In f.Controls
            If Left$(c.Tag, 6) = "LOCAL:" Then c.Visible = (CLng(Mid$(c.Tag, 7)) = section)
            If c.Name Like "uxSection#" Then
                c.BackColor = UXColor(IIf(c.Name = "uxSection" & section, "Active", "Neutral"))
                c.FontBold = (c.Name = "uxSection" & section)
            End If
        Next
        f!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
        f!cmdSave.Enabled = Not locked: f!cboKind.Locked = locked
        If locked Then f!uxLocked.Caption = "Read only. Recover archived reports or reopen local approval to edit." Else f!uxLocked.Caption = ""
        If Not UXSurfaceSection(section, locked) Then Exit Function
        UXResizeEditor
        UXRefreshEditor
        UXSelectSection = True: Exit Function
    End If
    For Each c In f.Controls
        If Left$(c.Tag, 3) = "UX:" Then
            tags = Split(c.Tag, ":")
            c.Visible = (CLng(tags(1)) = section And CLng(tags(2)) = page)
            If CLng(tags(1)) = section And CLng(tags(2)) > maxPage Then maxPage = CLng(tags(2))
            If c.ControlType = acTextBox Or c.ControlType = acComboBox Or c.ControlType = acCheckBox Then c.Locked = locked Or tags(3) = "Guidance"
        End If
    Next
    f!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
    f!uxPage.Caption = "Page " & (page + 1) & " of " & (maxPage + 1)
    f!uxPrevious.Enabled = (page > 0): f!uxNext.Enabled = (page < maxPage)
    f!cmdSave.Enabled = Not locked
    f!cboKind.Locked = locked
    If locked Then f!uxLocked.Caption = "Read only. Reopen local approval or recover the archived report to edit." Else f!uxLocked.Caption = ""
    UXRefreshEditor
    UXSelectSection = True
End Function
' Return a changed report type to its first section so form-specific fields and captions refresh.
Public Function UXEditorKindChanged() As Boolean
    Call EditorTypeChanged()
    UXEditorKindChanged = UXSelectSection(currentSection)
End Function
' Move within the current section's paging layout.
Public Function UXEditorPage(ByVal delta As Long) As Boolean
    UXEditorPage = UXSelectSection(currentSection, currentPage + delta)
End Function
' Map legacy checkbox state into canonical saved-field values, then invalidate and refresh that field.
Public Function UXFieldChanged(ByVal controlName As String, ByVal key As String) As Boolean
    Dim f As Access.Form, value As Variant
    Set f = Forms!frmEditor
    value = f.Controls(controlName).Value
    If key = "ObservationStatus" Then
        If IsNull(value) Then f!cboObservationStatus = Null Else If value Then f!cboObservationStatus = "NOB" Else f!cboObservationStatus = "Observed"
        If Nz(f!cboObservationStatus, "") = "NOB" Then UXSetWholeNOB Else UXClearWholeNOB
    Else
        If IsNull(value) Then f.Controls("txt" & key) = Null Else If value Then f.Controls("txt" & key) = "Yes" Else f.Controls("txt" & key) = "No"
    End If
    Call UXEditorChanged(): UXRefreshEditor key: UXFieldChanged = True
End Function
' Refresh pending header/ITA and field presentation. Membership and saved SGA are
' read only when their context is invalidated or selected group/report changes; a field
' key limits pending ITA work. Saved SGA never includes unsaved trait edits.
Public Function UXRefreshEditor(Optional ByVal changedKey As String = "") As Boolean
    Dim f As Access.Form, id As Long, r As DAO.Recordset, db As DAO.Database, groups As Long, groupID As Long, total As Double, count As Long, average As Variant, c As Access.Control, key As String, raw As Variant, memberLabel As String, priorAlias As String
    Set f = Forms!frmEditor
    If Not UXHasControl(f, "uxDirty") Then Exit Function
    f!uxDirty.Caption = IIf(EditorHasPending(), "Unsaved changes", "No pending changes")
    priorAlias = f!uxAlias.Caption
    memberLabel = UXIdentity(f!txtOfficialName, f!txtMember)
    f!uxAlias.Caption = ""
    If Len(Trim$(Nz(f!txtMember, ""))) > 0 And Trim$(Nz(f!txtMember, "")) <> Trim$(Nz(f!txtOfficialName, "")) And Len(Trim$(Nz(f!txtOfficialName, ""))) > 0 Then f!uxAlias.Caption = "Local name: " & Replace(f!txtMember, "&", "&&")
    f!uxReport.Caption = Replace(memberLabel, "&", "&&") & vbCrLf & Nz(f!cboKind, "") & " | " & Nz(f!txtRankRate, "Grade / rate") & " | " & Nz(f!txtStart, "Start date") & " to " & Nz(f!txtEnd, "End date")
    If changedKey = "" Or IsTraitKey(changedKey) Or changedKey = "ObservationStatus" Then
    f!uxITA.Caption = "ITA - pending: " & TraitSummary(TraitsFromEditor())
    If Nz(f!cboObservationStatus, "") = "NOB" Then f!uxITA.Caption = "ITA - pending: NOB (whole report)"
    End If
    id = Nz(f!txtReportID, 0)
    If editorContextDirty Or editorContextReport <> id Or editorContextGroup <> Nz(f!uxGroup, 0) Then
    editorGroupRefreshes = editorGroupRefreshes + 1
    id = Nz(f!txtReportID, 0): Set db = CurrentDb
    'Every active group of this report type; groups it is not in offer a move.
    f!uxGroup.RowSource = "SELECT g.GroupID,g.Title & IIf(DCount('*','GroupMember','GroupID=' & g.GroupID & ' AND ReportID=" & id & "')>0,'',' (move here)') AS Label FROM SummaryGroup AS g WHERE g.IsArchived=False AND g.FormKind=" & QuoteSQL(Nz(f!cboKind, "")) & " ORDER BY g.Title"
    f!uxGroup.Requery
    Set r = db.OpenRecordset("SELECT GroupID FROM GroupMember WHERE ReportID=" & id, dbOpenSnapshot)
    Do While Not r.EOF: groups = groups + 1: groupID = r!GroupID: r.MoveNext: Loop: r.Close
    If groups = 1 And IsNull(f!uxGroup) Then f!uxGroup = groupID
    If groups = 0 Then f!uxGroup = Null
    If Nz(f!uxGroup, 0) > 0 Then
        Set r = db.OpenRecordset("SELECT r.* FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & f!uxGroup & " AND r.IsArchived=False", dbOpenSnapshot)
        Do While Not r.EOF
            If Nz(r!ObservationStatus, "") <> "NOB" Then
                For Each average In TraitKeys()
                    raw = r.Fields(CStr(average)).Value
                    Select Case Nz(raw, "")
                        Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(raw): count = count + 1
                    End Select
                Next
            End If
            r.MoveNext
        Loop
        r.Close
        If count > 0 Then f!uxSGA.Caption = "Summary Group Average - saved values: " & Format$(total / count, "0.00") & " (" & count & " graded traits; arithmetic only)" Else f!uxSGA.Caption = "Summary Group Average: unavailable (no graded traits)"
    Else
        f!uxSGA.Caption = "Summary Group Average: " & IIf(groups > 1, "choose the explicit group context", "not assigned to a group")
    End If
    UXEditorSiblingState
    editorContextReport = id: editorContextGroup = Nz(f!uxGroup, 0): editorContextDirty = False
    End If
    For Each c In f.Controls
        If Left$(c.Name, 5) = "uxChk" Then
            key = Mid$(c.Name, 6)
            If key = "ObservationStatus" Then raw = f!cboObservationStatus.Value Else raw = f.Controls("txt" & key).Value
            'An unanswered box shows blank, as in NAVFIT98A; the buffer keeps it unanswered.
            If IsNull(raw) Or Len(Nz(raw, "")) = 0 Then c.Value = False Else If key = "ObservationStatus" Then c.Value = (raw = "NOB") Else c.Value = (raw = "Yes")
        End If
    Next
    If UXSurfaceAvailable() Then
        If Not UXSurfaceRefresh() Then Exit Function
    End If
    If priorAlias <> f!uxAlias.Caption Then UXResizeEditor
    'First paint does not wait for full policy checks. The existing one-second
    'dirty timer runs them immediately after display; field exits still validate.
    If Not editorStarting Then UXLivePaint
    UXRefreshEditor = True
End Function
' Keep pending values visible and explain how to undo them after a failed save.
Public Function UXShowSaveError(ByVal message As String, Optional ByVal committed As Boolean = False) As Boolean
    Dim f As Access.Form, primary As String, detail As String, separator As Long
    Set f = Forms!frmEditor
    If Not UXHasControl(f, "uxError") Then Exit Function
    message = SwiftRedact(message): separator = InStr(message, vbCrLf)
    If separator > 0 Then primary = Left$(message, separator - 1): detail = Mid$(message, separator + 2) Else primary = message: detail = message
    'Keep the existing fixed-height status readable; fuller redacted context is
    'available in its tooltip and SwiftLastFailure(), without resizing the editor.
    If committed Then
        f!uxError.Caption = "Saved; display refresh failed. Reopen the report. " & Left$(primary, 65)
        f!uxDirty.Caption = "Saved - reopen report to refresh display"
    Else
        f!uxError.Caption = "Not saved: " & Left$(primary, 80) & " Pending edits retained."
        f!uxDirty.Caption = "Not saved - pending edits retained"
    End If
    f!uxError.ControlTipText = Left$(detail, 255)
    UXShowSaveError = True
End Function
' Validate and flush the visible surface before the normal atomic draft save;
' refresh saved context afterward and retain the editor buffer on failure.
Public Function UXSaveEditor() As Boolean
    On Error GoTo Failed
    Dim f As Access.Form, saved As Boolean, context As Object
    Set f = Forms!frmEditor
    f!uxError.Caption = ""
    If UXSurfaceAvailable() Then
        If Not UXSurfaceCheckAll() Then
            f!uxDirty.Caption = "Not saved - correct invalid entries; pending edits retained"
            Exit Function
        End If
    End If
    saved = SaveEditor(): UXSaveEditor = saved
    If saved Then UXForgetWholeNOB
    If Not UXRefreshEditor() Then
        'The surface reported its own failure before the save outcome was known.
        Set context = SwiftLastFailureContext()
        If Not context Is Nothing Then
            If saved Then context("Transaction") = "Committed"
            UXShowSaveError SwiftFailureText(context), saved
        End If
    End If
    'A saved change over a later reviewer's edit is allowed but flagged (F01).
    Dim notice As String
    notice = RouteSaveNotice()
    If saved And Len(notice) > 0 Then f!uxError.Caption = Left$("Saved. " & notice & " They will see it flagged when the report reaches them.", 255)
    Exit Function
Failed:
    Set context = SwiftFailure("modUX", "UXSaveEditor", "Capture/refresh editor", Err.Number, Err.Description, Err.Source, IIf(saved, "Committed", "Not started"))
    UXShowSaveError SwiftFailureText(context), saved
    UXSaveEditor = saved
End Function
' Delegate summary-group directory loading to its current list adapter.
Public Function UXLoadGroups() As Boolean
    UXLoadGroups = UXGroupList()
End Function
' All group rosters use the same Members surface. prepareOnly retains a hidden
' adapter for legacy Review/History callers, never an alternate empty roster.
Public Function UXOpenGroup(Optional ByVal prepareOnly As Boolean = False) As Boolean
    On Error GoTo Failed
    groupContext = Nz(Forms!frmUXGroups!lstGroups, 0)
    If groupContext = 0 Then MsgBox "Choose a summary group.", vbInformation: Exit Function
    If Not UXLeaveEditor() Then Exit Function
    If Not UXLeaveGroup() Then Exit Function
    If Not prepareOnly Then UXOpenGroup = UXShowGroupMembers(groupContext): Exit Function
    DoCmd.Echo False
    If Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
    Forms!frmMajorWorkbench!cboRun = Null
    Forms!frmMajorWorkbench!cboGroup = groupContext
    Forms!frmMajorWorkbench.Visible = False
    DoCmd.Echo True
    UXOpenGroup = True: Exit Function
Failed:
    Dim failure As Object
    Set failure = SwiftFailure("modUX", "UXOpenGroup", "Open group context", Err.Number, Err.Description, IIf(Len(Err.Source) = 0, "modUX.UXOpenGroup", Err.Source))
    On Error Resume Next
    DoCmd.Echo True
    If Err.Number <> 0 Then SwiftCleanupError failure, "restore screen painting", Err.Number
    On Error GoTo 0
    Err.Raise CLng(failure("Number")), CStr(failure("Source")), CStr(failure("Description"))
End Function
' Show the retained group tools for one mode and refresh their focused group context.
Public Function UXWorkbenchMode(ByVal mode As String) As Boolean
    Dim f As Access.Form, c As Access.Control, tags As Variant
    Set f = Forms!frmMajorWorkbench
    If mode = "Members" And Nz(f!cboGroup, 0) > 0 Then
        UXWorkbenchMode = UXShowGroupMembers(CLng(f!cboGroup)): Exit Function
    End If
    f!uxMode = mode
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "WBX:" Then c.Visible = (InStr(1, ":" & Mid$(c.Tag, 5) & ":", ":" & mode & ":", vbBinaryCompare) > 0)
    Next
    If mode = "Members" Or mode = "Comparison" Or mode = "Defaults" Then f.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & Nz(f!cboGroup, 0)), "Summary group")
    UXFocusedGroupMode mode
    UXWorkbenchMode = True
End Function
' Populate available/member selections for explicit group membership editing.
Public Function UXChooseMembers() As Boolean
    Dim f As Access.Form, i As Long, id As Long, kind As String
    groupContext = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If groupContext = 0 Then MsgBox "Create or choose a group first.": Exit Function
    kind = Nz(MajorScalar("SELECT FormKind FROM SummaryGroup WHERE GroupID=" & groupContext), "")
    DoCmd.OpenForm "frmUXMembers": Set f = Forms!frmUXMembers
    'Unnamed drafts look alike, so each row also shows its report number, period, state and groups.
    f!lstReports.RowSource = "SELECT ReportID,UXIdentity(OfficialName,MemberName) & ' (report ' & ReportID & ')' AS Member,RankRate,Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No period') & ' | ' & UXReportState(ReportID) & ' | ' & UXTableGroups(ReportID) AS Details FROM Report WHERE IsArchived=False AND FormKind=" & QuoteSQL(kind) & " ORDER BY UXIdentity(OfficialName,MemberName),ReportID"
    f!lstReports.Requery
    For i = 0 To f!lstReports.ListCount - 1
        id = Nz(f!lstReports.Column(0, i), 0)
        f!lstReports.Selected(i) = (DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) > 0)
    Next
    GroupMembersStart groupContext
    UXChooseMembers = True
End Function
' Apply selected membership changes and return to the group member view.
Public Function UXSaveMembers() As Boolean
    Dim f As Access.Form, i As Long, id As Long, grade As String, w As DAO.Workspace, started As Boolean
    On Error GoTo Failed
    Set f = Forms!frmUXMembers: grade = Nz(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupContext), "")
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    For i = 0 To f!lstReports.ListCount - 1
        id = CLng(f!lstReports.Column(0, i))
        If f!lstReports.Selected(i) Then
            If DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) = 0 Then AddGroupMember groupContext, id, grade, Nz(f!txtNote, "")
        Else
            If DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) > 0 Then RemoveGroupMember groupContext, id
        End If
    Next
    w.CommitTrans: started = False: DoCmd.Close acForm, "frmUXMembers", acSaveNo
    UXSaveMembers = UXShowGroupMembers(groupContext, True): Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUX", "UXSaveMembers", "Execute", Err.number, Err.description, Err.source, "Unknown")
    If started Then SwiftRollback w, swiftPrimary
    Forms!frmUXMembers!lblStatus.caption = CStr(swiftPrimary("Description"))
End Function
' Open live Members or an explicitly requested historical comparison according to
' the workbench mode; reuse an existing window rather than replacing its pending edits.
Public Function UXOpenComparison() As Boolean
    Dim f As Access.Form, id As Long, snapshot As Long, title As String
    On Error GoTo Failed
    Set f = Forms!frmMajorWorkbench
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        Forms!frmUXComparison.SetFocus
        f!lblStatus.Caption = "A comparison is already open. Finish or close it before opening another group comparison."
        Exit Function
    End If
    If Nz(f!uxMode, "Members") = "Comparison" Then
        title = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & Nz(f!cboGroup, 0)), "Comparison")
        id = ReuseGroupComparison(Nz(f!cboGroup, 0), title & " comparison"): f!cboMeeting = id
        DoCmd.OpenForm "frmUXComparison", , , , , , CStr(id)
    Else
        DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & Nz(f!cboGroup, 0)
    End If
    UXGroupTabState "frmUXComparison", IIf(Nz(f!uxMode, "Members") = "Members", 0, 3)
    UXOpenComparison = True: Exit Function
Failed:
    Forms!frmMajorWorkbench!lblStatus.Caption = Err.Description
End Function
' Resolve live-group or historical-set context and load the editable grid plus one
' proposal preview; no duplicate hidden roster query is required.
Public Function UXLoadComparison() As Boolean
    Dim f As Access.Form, block As Long, sql As String
    Set f = Forms!frmUXComparison
    If Left$(Nz(f.OpenArgs, ""), 2) = "G:" Then
        f!txtMeeting = 0: f!uxGroupContext = CLng(Mid$(f.OpenArgs, 3))
        f!lblHeading.Caption = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & f!uxGroupContext), "Group") & " members"
    Else
        f!txtMeeting = CLng(f.OpenArgs)
        f!lblHeading.Caption = Nz(MajorScalar("SELECT Title FROM ComparisonSet WHERE SetID=" & f!txtMeeting), "Comparison")
        f!uxGroupContext = Nz(MajorScalar("SELECT g.GroupID FROM GroupSnapshot AS g INNER JOIN ComparisonSet AS c ON g.SnapshotID=c.GroupSnapshotID WHERE c.SetID=" & f!txtMeeting), 0)
    End If
    UXRefreshComparisonPreview
    UXGridLoad
    If UXHasControl(f, "uxMarksPage") Then UXComparisonPage "Marks"
    UXLoadComparison = True
End Function
' Populate the retained proposal adapter from the active grid row and its saved
' revision; the obsolete hidden roster selector is no longer required.
Public Function UXSelectComparisonMember() As Boolean
    Dim f As Access.Form, block As Long, id As Long, locked As Boolean
    Set f = Forms!frmUXComparison: id = Nz(f!uxTraitGrid.Form!ReportID, 0): If id = 0 Then Exit Function
    For block = 33 To 39: f.Controls("cboTrait" & block) = MajorRevisionValue(f!uxTraitGrid.Form!RevisionID, "Trait" & block): Next
    f!cboPromotion = MajorRevisionValue(f!uxTraitGrid.Form!RevisionID, "PromotionRecommendation")
    locked = Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & id), True) Or Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & id), "") <> "Draft"
    f!cmdStage.Enabled = Not locked
    For block = 33 To 39: f.Controls("cboTrait" & block).Locked = locked: Next
    f!cboPromotion.Locked = locked
    f!uxMember.Caption = Nz(f!uxTraitGrid.Form!Member, "") & IIf(locked, " - locked; reopen in the report workspace to edit", " - proposed edits; stage and review before Apply")
    Call UXComparisonAverage: UXSelectComparisonMember = True
End Function
' Display the proposed trait average from the retained proposal controls.
Public Function UXComparisonAverage() As Boolean
    Dim values As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null)
    For i = 0 To 6: values(i) = Forms!frmUXComparison.Controls("cboTrait" & (33 + i)): Next
    Forms!frmUXComparison!uxITA.Caption = "Proposed ITA: " & TraitSummary(values)
    UXComparisonAverage = True
End Function
' Stage changes against the selected row's saved revision. Materialize a durable
' comparison only for explicit proposal work; do not modify saved reports during staging.
Public Function UXStageComparison() As Boolean
    Dim f As Access.Form, id As Long, revision As Long, block As Long, before As Variant, value As Variant, count As Long, w As DAO.Workspace, started As Boolean
    On Error GoTo Failed
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, "modUX.UXStageComparison", "Retry or discard pending cells before staging proposals."
    Set f = Forms!frmUXComparison: id = Nz(f!uxTraitGrid.Form!ReportID, 0)
    If id = 0 Then Err.Raise vbObjectError + 4700, "modUX.UXStageComparison", "Choose a member in the comparison grid."
    revision = Nz(f!uxTraitGrid.Form!RevisionID, 0)
    If CurrentRevision(id) <> revision Then Err.Raise vbObjectError + 4810, "modUX.UXStageComparison", "Saved report changed. Reopen Members before staging proposals."
    If Nz(f!txtMeeting, 0) = 0 Then
        f!txtMeeting = ReuseGroupComparison(f!uxGroupContext, "Group proposals")
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    For block = 33 To 39
        before = MajorRevisionValue(revision, "Trait" & block): value = f.Controls("cboTrait" & block).Value
        If Not SameFieldValue(before, value) Then ProposeMeetingField f!txtMeeting, id, "Trait" & block, value: count = count + 1
    Next
    before = MajorRevisionValue(revision, "PromotionRecommendation"): value = f!cboPromotion.Value
    If Not SameFieldValue(before, value) Then ProposeMeetingField f!txtMeeting, id, "PromotionRecommendation", value: count = count + 1
    w.CommitTrans: started = False
Call UXGridLoad: UXRefreshComparisonPreview
    f!lblStatus.Caption = count & " proposals staged. Saved reports are unchanged. Select one proposal per field below, then Apply."
    If UXHasControl(f, "uxPreviewPage") And count > 0 Then UXComparisonPage "Preview"
    UXStageComparison = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUX", "UXStageComparison", "Execute", Err.number, Err.description, Err.source, "Unknown")
    If started Then SwiftRollback w, swiftPrimary
    Forms!frmUXComparison!lblStatus.caption = CStr(swiftPrimary("Description"))
End Function
' Apply selected reviewed proposals atomically, then show results and treat that
' comparison as historical; pending direct-edit cells must be resolved first.
Public Function UXApplyComparison() As Boolean
    Dim item As Variant, csv As String
    On Error GoTo Failed
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, "modUX.UXApplyComparison", "Retry or discard pending cells before applying proposals."
    For Each item In Forms!frmUXComparison!lstProposals.ItemsSelected
        If Len(csv) > 0 Then csv = csv & ","
        csv = csv & Forms!frmUXComparison!lstProposals.ItemData(item)
    Next
    ApplyMeetingProposals csv
    Forms!frmUXComparison!cmdStage.Enabled = False
    Call UXGridLoad: Call UXRefreshComparisonPreview(True)
    UXComparisonPage "Preview"
    Forms!frmUXComparison!cmdBack.SetFocus
    Forms!frmUXComparison!cmdApply.Enabled = False
    Forms!frmUXComparison!lblStatus.Caption = "Selected proposals applied atomically. This comparison is now historical; reopen the group comparison for further edits."
    UXApplyComparison = True: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = Err.Description & " No partial proposal batch was applied."
End Function
' Resolve whether review should use saved or pending values before opening saved-revision review.
Public Function UXEditorReview(ByVal mode As String) As Boolean
 UXEditorReview = UXToolsChoose("Review")
End Function
' Seed the contextual shared-block/default editor from saved values, tracking mixed
' fields/defaults before edits are confirmed and applied recursively.
Public Function UXStartGroupBulk(Optional ByVal defaultsMode As Boolean = False, Optional ByVal globalMode As Boolean = False, Optional ByVal requestedGroup As Long = -1) As Boolean
    Dim r As DAO.Recordset, key As Variant, values As Object, f As Access.Form, seen As Boolean, base As Object, db As DAO.Database
    If globalMode Then
        groupContext = 0
    ElseIf requestedGroup >= 0 Then
        groupContext = requestedGroup
    ElseIf CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then
        groupContext = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    Else
        Exit Function
    End If
    If groupContext = 0 And Not globalMode Then MsgBox "Select a group first.": Exit Function
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        If Not UXGridDrain() Then Exit Function
    End If
    Set bulkMixed = CreateObject("Scripting.Dictionary")
    If groupContext > 0 Then
        If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupContext), True) Then Exit Function
    End If
    ' A workspace tab can leave this shared form open. Finish its active control
    ' under the old confirmation context, then reopen with fresh OpenArgs/buffers.
    ' Closing an unbound form does not discard confirmed report/default changes.
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        Forms!frmUXBulk.SetFocus
        Forms!frmUXBulk!cmdBack.SetFocus
        DoCmd.Close acForm, "frmUXBulk", acSaveNo
    End If
    bulkDefaults = defaultsMode: bulkDefaultStamp = UXDefaultsStamp(groupContext)
    bulkTargetStamp = UXBulkContextStamp(groupContext)
    bulkCSV = "": Set bulkChanges = CreateObject("Scripting.Dictionary"): Set db = CurrentDb
    Set bulkDefaultResets = CreateObject("Scripting.Dictionary")
    Set r = db.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupContext & " ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        If Len(bulkCSV) > 0 Then bulkCSV = bulkCSV & ","
        bulkCSV = bulkCSV & r!ReportID
        Set values = SavedFields(r!ReportID)
        If Not seen Then
            Set base = values: seen = True
        Else
            For Each key In base.Keys
                If Not SameFieldValue(base(key), values(key)) Then base(key) = Null: bulkMixed(key) = True
            Next
        End If
        r.MoveNext
    Loop
    r.Close
    If Not seen And Not defaultsMode Then MsgBox "Add group members first.": Exit Function
    If base Is Nothing Then Set base = CreateObject("Scripting.Dictionary")
    If defaultsMode Then
        For Each key In UXSharedKeys(): base(key) = UXEffectiveDefault(groupContext, CStr(key)): Next
    End If
    DoCmd.OpenForm "frmUXBulk", , , , , , CStr(groupContext)
    Set f = Forms!frmUXBulk
    f!uxGroupContext = groupContext
    UXGroupTabState f.Name, IIf(defaultsMode, 1, 0)
    UXInvalidateBulkPreview
    f!uxKind = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & groupContext), "EVAL")
    f!lblHeading.Caption = IIf(globalMode, "Global defaults", IIf(defaultsMode, "Group defaults", "Shared report blocks"))

    For Each key In UXSharedKeys()
        If UXHasControl(f, "uxValue" & key) Then
            If f.Controls("uxValue" & key).ControlType = acCheckBox Then
                If IsNull(base(key)) Or Len(Nz(base(key), "")) = 0 Then
                    f.Controls("uxValue" & key) = Null
                ElseIf key = "ObservationStatus" Then
                    f.Controls("uxValue" & key) = (base(key) = "NOB")
                Else
                    f.Controls("uxValue" & key) = (base(key) = "Yes")
                End If
            Else
                f.Controls("uxValue" & key) = base(key)
            End If
        End If
    Next
    f!uxGroupTitle.Caption = IIf(globalMode, "All active summary groups", Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & groupContext), "Summary group"))
    UXBlockRefresh
    f!uxStatus.Caption = "Edit a field to confirm and apply it."
    Call UXBeginImmediateBlocks(groupContext, defaultsMode, bulkCSV)
    UXBlockMode "Edit"
    Call UXBulkSection(0)
    UXStartGroupBulk = True
End Function
' Confirm and apply one explicit shared-field edit.
Public Function UXBulkChanged(ByVal key As String) As Boolean
    UXBulkChanged = UXConfirmBlockChange(key)
End Function
' Compatibility cleanup: direct confirmed edits no longer stage a preview.
Public Function UXInvalidateBulkPreview(Optional ByVal clearPreview As Boolean = True) As Boolean
    bulkBatch = 0: bulkPreviewReason = "": bulkPreviewStamp = ""
    UXInvalidateBulkPreview = True
End Function
' Expose a staged block value to Access preview expressions.
Public Function UXBulkPending(ByVal key As String) As Variant
    UXBulkPending = bulkDefaultPreview(key)
End Function
' Expose the saved effective default captured before preview to Access expressions.
Public Function UXBulkBefore(ByVal key As String) As Variant
    UXBulkBefore = bulkDefaultBefore(key)
End Function
' Navigate the shared-block editor's section without applying changes.
Public Function UXBulkSection(ByVal section As Long, Optional ByVal page As Long = 0) As Boolean
    If section = 1 Then section = 0
    Forms!frmUXBulk!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
    UXBlockMode "Edit": UXBulkSection = UXBlockLayout(section)
End Function
' Validate edited blocks and capture complete before/after values and context stamps.
' Preview remains separate from the later atomic apply.
Public Function UXPreviewBulk() As Boolean
    'Retired UI entry point: field edits now confirm and commit immediately.
End Function
' Reject a changed preview/context, then commit defaults or the report batch through
' the appropriate transactional service; stale previews are invalidated before writes.
Public Function UXApplyBulk() As Boolean
    'Retired UI entry point: field edits now confirm and commit immediately.
End Function
' Return whether the shared-block editor currently edits defaults rather than saved reports.
Public Function UXEditingDefaults() As Boolean
    UXEditingDefaults = bulkDefaults
End Function
' Stage removal of a group/global default override so parent inheritance can apply again.
Public Function UXResetDefault(ByVal key As String) As Boolean
    'Retired UI entry point; edit the field directly to replace or blank its value.
End Function
' Show complete values for the selected preview item, outside its condensed row.
Public Function UXBulkDetail() As Boolean
    'Retired UI entry point: field edits now confirm and commit immediately.
End Function
' Update workspace/contributor headers only when their text changes and resize the active workspace.
Public Function UXLoadContext() As Boolean
    Dim f As Access.Form, active As Access.Form, context As String
    context = Replace(Nz(DLookup("CommandName", "Workspace"), "Swift NAVFIT") & " | Contributor: " & ActiveContributorCaption(), "&", "&&")
    For Each f In Forms
        If UXHasControl(f, "uxContext") Then
            If f!uxContext.Caption <> context Then f!uxContext.Caption = context
        End If
    Next
    'Only the visible workspace needs geometry; unchanged dimensions are cached.
    On Error Resume Next
    Set active = Screen.ActiveForm
    On Error GoTo 0
    If Not active Is Nothing Then
        If active.Name <> "frmEditor" And UXHasControl(active, "uxContext") Then UXResizeWorkspace active.Name
    End If
    UXLoadContext = True
End Function

' Close a utility form without saving design changes.
Public Function UXCloseScreen(ByVal formName As String) As Boolean
    DoCmd.Close acForm, formName, acSaveNo
    UXCloseScreen = True
End Function
' Return resource tools to Resources; snippets use their report-return adapter instead.
Public Function UXCloseResource(ByVal formName As String) As Boolean
    If formName = "frmSnippets" Then UXCloseResource = UXReturnSnippets(): Exit Function
    DoCmd.Close acForm, formName, acSaveNo
    DoCmd.OpenForm "frmUXResources"
    UXCloseResource = True
End Function

' Load retained group-tool selectors, skipping the Members refresh when opened as a context-only host.
Public Function UXLoadWorkbench() As Boolean
    If Not LoadMajorWorkbench() Then Exit Function
    If Nz(Forms!frmMajorWorkbench.OpenArgs, "") <> "ContextOnly" Then Call UXWorkbenchMode("Members")
    UXLoadWorkbench = True
End Function


' Refresh workflow locks only when the affected report is the open editor.
Public Function UXRefreshEditorWorkflow(ByVal reportID As Long) As Boolean
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then UXRefreshEditorWorkflow = True: Exit Function
    If Nz(Forms!frmEditor!txtReportID, 0) <> reportID Then UXRefreshEditorWorkflow = True: Exit Function
    Call UXFlushSurface
    UXRefreshEditorWorkflow = UXSelectSection(currentSection, currentPage)
End Function

' Detect a live bulk preview or staged field/default edits before allowing a group-page exit.
Public Function UXGroupBulkPending() As Boolean
    'There is no staged defaults/bulk preview to hold navigation open.
End Function
