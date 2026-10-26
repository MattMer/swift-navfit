Attribute VB_Name = "modUXValidation"
Option Compare Binary
Option Explicit
Private issues As Collection
Private priorSignature As String
Private nextCheckpoint As Date
Private busy As Boolean
Private pageIndex As Long
Private liveSignature As String
Private liveResult As Collection
Private validationDirty As Boolean
Private recoveryDirty As Boolean
Private validationDate As Date
Private paintCount As Long
' Escape a value for an Access value-list cell; flatten newlines for the list display.
Private Function Cell(ByVal value As Variant) As String
    Cell = Chr$(34) & Replace(Replace(Replace(Nz(value, ""), Chr$(34), Chr$(34) & Chr$(34)), vbCr, " "), vbLf, " ") & Chr$(34)
End Function
' Validate pending editor values and paint the issue summary. Cache findings by exact
' inputs and date; dirty/save state is independent of saved report revisions.
Public Function UXLivePaint() As Boolean
    Dim result As Collection, item As Object, invalid As Long, policy As Long, incomplete As Long, review As Long
    Dim pending As Object, key As Variant, signature As String
    If busy Or Not CurrentProject.AllForms("frmEditor").IsLoaded Then Exit Function
    On Error GoTo Failed
    paintCount = paintCount + 1
    busy = True: Set pending = EditorPendingFields()
    Call UXNarrativeStatus()
    signature = Nz(Forms!frmEditor!txtReportID, 0) & "|" & Date & "|"
    For Each key In pending.Keys
        signature = signature & IIf(IsNull(pending(key)), "N", "V") & Len(Nz(pending(key), "")) & ":" & Nz(pending(key), "")
    Next
    If liveResult Is Nothing Or signature <> liveSignature Then
        Set liveResult = ValidateLiveFields(pending, Nz(Forms!frmEditor!txtReportID, 0))
        liveSignature = signature
    End If
    Set result = liveResult
    For Each item In result
        Select Case item("Severity")
            Case "Invalid entry": invalid = invalid + 1
            Case "Policy issue": policy = policy + 1
            Case "Incomplete draft": incomplete = incomplete + 1
            Case "Needs review": review = review + 1
        End Select
    Next
    Forms!frmEditor!uxDirty.Caption = IIf(EditorHasPending(), "Unsaved changes", "No pending changes") & " | " & result.count & " issues"
    Forms!frmEditor!uxDirty.ControlTipText = "Open validation: " & invalid & " invalid entries, " & policy & " policy issues, " & incomplete & " incomplete, " & review & " need review. Pending values; no saved revision created."
    If EditorHasPending() Then recoveryDirty = True
    validationDirty = False: validationDate = Date
    UXLivePaint = True
Finished:
    UXApplyVisual Forms!frmEditor
    busy = False: Exit Function
Failed:
    Forms!frmEditor!uxDirty.Caption = "Validation unavailable - open issues"
    Resume Finished
End Function
' Expose the validation-paint count for diagnostics without querying the database.
Public Function UXLivePaintCount() As Long
    UXLivePaintCount = paintCount
End Function
' Mark validation and recovery dirty after a local edit; defer typing checks to the timer.
Public Function UXEditorChanged() As Boolean
    validationDirty = True: recoveryDirty = True
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        Forms!frmEditor!lblRecovery.Caption = "Recovery: latest edits are not checkpointed yet. Save draft or keep this editor open."
        If Forms!frmEditor!cmdSave.Enabled Then
            Forms!frmEditor!uxLocked.Caption = "Next: Save draft; review uses saved values."
            Dim action As Variant
            For Each action In Array("uxReviewSetup", "uxReviewAccept", "uxReviewReturn", "uxReviewApproval")
                If UXHasControl(Forms!frmEditor, CStr(action)) Then
                    Forms!frmEditor.Controls(action).FontBold = False: Forms!frmEditor.Controls(action).BackColor = UXColor("Neutral")
                End If
            Next
            Forms!frmEditor!cmdSave.FontBold = True: Forms!frmEditor!cmdSave.BackColor = UXColor("Primary")
        End If
    End If
    UXEditorChanged = True
End Function
' Reset validation and checkpoint scheduling for a new editor session.
Public Function UXLiveReset() As Boolean
    validationDirty = True: recoveryDirty = False: nextCheckpoint = DateAdd("s", 30, Now())
    liveSignature = "": Set liveResult = Nothing
    UXLiveReset = True
End Function
' Handle textbox OnChange without validating on every keystroke. Field exit and the
' dirty timer flush the checks; active text remains available to recovery.
Public Function UXLiveTyped() As Boolean
    'First keystroke in a signed report: ask to reopen, or undo the keystroke.
    If Not UXSignedEditGuard() Then
        On Error Resume Next
        Screen.ActiveControl.Undo
        Exit Function
    End If
    Call UXEditorChanged()
    UXLiveTyped = True
End Function
' Process dirty validation and due recovery work only. Failed checkpoints retain the
' dirty flag so a later tick retries; an unchanged editor does not scan its fields.
Public Function UXLiveTimer() As Boolean
    If validationDirty Or validationDate <> Date Then Call UXLivePaint()
    If nextCheckpoint = 0 Then nextCheckpoint = DateAdd("s", 30, Now())
    If Now() >= nextCheckpoint Then
        If recoveryDirty Then
            If CheckpointEditor() Then recoveryDirty = False
        End If
        nextCheckpoint = DateAdd("s", 30, Now())
    End If
    UXToolsTimer
    UXLiveTimer = True
End Function
' Open report or group findings using an explicit context; reset paging and selection.
Public Function UXOpenValidation(Optional ByVal groupMode As Boolean = False, Optional ByVal groupID As Long = 0) As Boolean
    Dim context As String
    If groupMode Then
        If groupID = 0 Then groupID = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
        context = "G:" & groupID
    Else
        context = "R:" & Nz(Forms!frmEditor!txtReportID, 0)
    End If
    If CurrentProject.AllForms("frmUXValidation").IsLoaded Then DoCmd.Close acForm, "frmUXValidation", acSaveNo
    priorSignature = "": pageIndex = 0: Set issues = Nothing
    DoCmd.OpenForm "frmUXValidation", , , , , , context
    UXOpenValidation = True
End Function
' Export pages pass an explicit saved report/group context and expose every finding.
Public Function UXExportReadiness(ByVal groupMode As Boolean) As Boolean
    Dim id As Long, context As String
    With Forms!frmNAVFITExchange
        If groupMode Then id = Nz(!cboGroup, 0): context = "G:" Else id = Nz(!cboReport, 0): context = "R:"
        If id = 0 Then !txtStatus = "Choose a report or summary group first.": Exit Function
    End With
    If CurrentProject.AllForms("frmUXValidation").IsLoaded Then DoCmd.Close acForm, "frmUXValidation", acSaveNo
    priorSignature = "": pageIndex = 0: Set issues = Nothing
    DoCmd.OpenForm "frmUXValidation", , , , , , context & id
    UXExportReadiness = True
End Function
' Evaluate current findings, including open pending values, and show 50 per page.
' The display signature suppresses redundant list repainting, not the underlying checks.
Public Function UXLoadValidation() As Boolean
    Dim f As Access.Form, item As Object, result As Collection, id As Long, signature As String, rows As String, i As Long, selection As String, identity As String, selected As Long
    On Error GoTo Failed
    Set f = Forms!frmUXValidation
    id = CLng(Mid$(Nz(f.OpenArgs, "R:0"), 3))
    If Left$(Nz(f.OpenArgs, "R:0"), 2) = "G:" Then
        Set result = ValidateLiveGroup(id)
        f!lblBasis.Caption = "All group members: saved values, plus the open member's pending values. Read-only checks; no approval or saved validation run. Verify completeness and exceptions."
    Else
        If CurrentProject.AllForms("frmEditor").IsLoaded Then
            If Nz(Forms!frmEditor!txtReportID, 0) = id Then Set result = ValidateLiveFields(EditorPendingFields(), id)
        End If
        If result Is Nothing And id > 0 Then Set result = ValidateLiveFields(SavedFields(id), id)
        If result Is Nothing Then Set result = New Collection
        f!lblBasis.Caption = "Current report values, including pending edits when open. Incomplete drafts and policy findings can be saved. Application entry checks remain separate."
    End If
    selection = ""
    If Not issues Is Nothing And Nz(f!lstIssues, 0) > 0 Then
        If CLng(f!lstIssues) <= issues.count Then
            Set item = issues(CLng(f!lstIssues)): selection = item("ReportID") & "|" & item("Key") & "|" & item("Code")
        End If
    End If
    For Each item In result
        i = i + 1: identity = item("ReportID") & "|" & item("Key") & "|" & item("Code")
        signature = signature & identity & item("Severity") & item("Message") & vbLf
        If identity = selection Then selected = i
        If i > pageIndex * 50 And i <= (pageIndex + 1) * 50 Then rows = rows & i & ";" & Cell(item("Member")) & ";" & Cell(item("Label")) & ";" & Cell(item("Severity")) & ";" & Cell(Left$(item("Message"), 180)) & ";"
    Next
    If pageIndex > 0 And i <= pageIndex * 50 Then pageIndex = 0: priorSignature = "": UXLoadValidation: Exit Function
    signature = Nz(f.OpenArgs, "") & "|" & pageIndex & "|" & signature
    If signature <> priorSignature Then
        Set issues = result: priorSignature = signature
        f!lstIssues.RowSourceType = "Value List": f!lstIssues.RowSource = "0;Member;Report block;Finding;Issue;" & rows
        If (selected <= pageIndex * 50 Or selected > (pageIndex + 1) * 50) And i > 0 Then selected = pageIndex * 50 + 1
        If selected > 0 Then f!lstIssues = selected Else f!lstIssues = Null
        UXValidationSelection
    End If
    f!lblCount.Caption = i & " findings from implemented checks. Coverage is limited; zero findings does not establish Navy readiness."
    f!cmdPrevious.Enabled = pageIndex > 0: f!cmdNext.Enabled = i > (pageIndex + 1) * 50
    f!cmdNext.ControlTipText = "Page " & (pageIndex + 1) & "; 50 findings per page. Full finding and citation appear below."
    UXLoadValidation = True: Exit Function
Failed:
    f!lblCount.Caption = "Validation failed: " & Err.Description
    f!cmdOpen.Enabled = False
End Function
' Show the selected finding and its policy citation; enable its field-navigation action.
Public Function UXValidationSelection() As Boolean
    
    Dim item As Object
    With Forms!frmUXValidation
        .cmdOpen.Enabled = False: .txtEvidence = "No findings from implemented checks."
        If issues Is Nothing Or Nz(.lstIssues, 0) = 0 Then Exit Function
        Set item = issues(CLng(.lstIssues))
        .txtEvidence = item("Severity") & ": " & item("Message") & vbCrLf & vbCrLf & item("Reference")
        .cmdOpen.Enabled = True
    End With
    UXValidationSelection = True
End Function
' Move the findings page and rebuild its visible rows, keeping page indices nonnegative.
Public Function UXValidationPage(ByVal offset As Long) As Boolean
    pageIndex = pageIndex + offset: If pageIndex < 0 Then pageIndex = 0
    UXValidationPage = UXLoadValidation()
End Function
' Navigate to the finding's report section and focus its field without losing another
' report's pending edits; failed navigation leaves the finding available.
Public Function UXValidationOpenField() As Boolean
    Dim item As Object, id As Long
    On Error GoTo Failed
    Set item = issues(CLng(Forms!frmUXValidation!lstIssues)): id = item("ReportID")
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) <> id Then
            If Not UXLeaveEditor() Then Exit Function
        End If
    End If
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    If UXFixedEnabled() Then UXFixedOpenField CStr(item("Key")) Else UXSelectSection UXFieldSection(item("Key"))
    Forms!frmEditor.SetFocus
    If UXSurfaceAvailable() Then
        If UXHasControl(Forms!frmEditor!uxFieldSurface.Form, UXFieldControl(item("Key"))) Then
            Forms!frmEditor!uxFieldSurface.SetFocus
            Forms!frmEditor!uxFieldSurface.Form.Controls(UXFieldControl(item("Key"))).SetFocus
        Else
            Forms!frmEditor.Controls(UXFieldControl(item("Key"))).SetFocus
        End If
    End If
    UXValidationOpenField = True: Exit Function
Failed:
    Forms!frmUXValidation!txtEvidence = Err.Description
End Function
' Release the findings cache and close its window.
Public Function UXValidationClose() As Boolean
    priorSignature = "": Set issues = Nothing
    DoCmd.Close acForm, "frmUXValidation", acSaveNo
    UXValidationClose = True
End Function
' Size the findings list and evidence panel for the available native form viewport.
Public Function UXValidationLayout() As Boolean
    Dim f As Access.Form, width As Long, height As Long
    On Error GoTo Finished
    Set f = Forms!frmUXValidation: width = f.InsideWidth - 800: height = f.InsideHeight
    If width > UX_MAX_FORM_WIDTH - 1000 Then width = UX_MAX_FORM_WIDTH - 1000
    If width < 4500 Then Exit Function
    f!lblHeading.Width = width - 750: f!lblBasis.Width = width: f!lblCount.Width = width - 2700
    f!cmdWindowHelp.Left = width - 100
    f!cmdPrevious.Left = width - 2200: f!cmdNext.Left = width - 850
    f!lstIssues.Width = width: f!txtEvidence.Width = width
    f!lblIssues.Width = width: f!lblEvidence.Width = width
    f!lstIssues.ColumnWidths = "0;" & CLng(width * 0.2) & ";" & CLng(width * 0.25) & ";" & CLng(width * 0.15) & ";" & CLng(width * 0.38)
    If height < 8500 Then height = 8500
    f!txtEvidence.Height = height - 6450
    f!cmdOpen.Top = height - 650: f!cmdBack.Top = height - 650
    f!cmdOpen.Width = width \ 2 - 100: f!cmdBack.Left = 500 + width \ 2: f!cmdBack.Width = width \ 2 - 100
    f.Width = width + 800: f.Section(acDetail).Height = height: f.ScrollBars = 2
    UXValidationLayout = True
Finished:
End Function
'Live diagnostics never commit fields, revisions, validation runs or rule packs.
