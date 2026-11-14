Attribute VB_Name = "modUXHelp"
Option Compare Database
Option Explicit
Public Function UXHelpText(ByVal topic As Long) As String
    Select Case topic
        Case 0: UXHelpText = "Use the synthetic demo to explore. Use the blank candidate for a separate workspace; complete the command and contributor setup. Keep existing workspaces intact. Reports, Summary groups and My Work are the main destinations. Local contributor profiles provide attribution, not authentication. To carry an existing workspace forward, open a new blank release and choose Import data from a previous Swift NAVFIT version on first-run setup or in Settings & Help. The seeded demo is for exploration, not a migration destination. Setup keeps your work in a data file, Swift NAVFIT data.accdb, beside this application (Change folder to pick another); this " & _
            "file stays the application, so a newer version can replace it. Open an existing data file uses one you already have; given a complete workspace from an earlier version it brings it into a new data file beside it, reading the earlier file only. If the data file cannot be opened at startup, a window says why and offers Find the data file, Create a new data file and Close Swift NAVFIT."
        Case 1: UXHelpText = "Reports > New report: choose EVAL, CHIEFEVAL or FITREP. Front Page and Back Page follow official block order with fixed field positions beside the main navigation. All seven performance traits are together on Front Page; this editing grouping does not change the printed form. Resize the viewing area without changing field geometry. Checkbox blocks remain explicit selections; SSNs retain leading zeros. Export-only structured name fields stay behind the visible full-name entry. Save draft atomically saves the current report. Page changes preserve pending values without creating revisions. Leaving a changed report saves automatically. Save failu" & _
            "res preserve pending values in Recovery when a checkpoint succeeds, allowing navigation without discarding invalid drafts. Open Settings & Help > Recover pending edits to resume them. Report tools contains Notes and replies; History remains separate. Comments retain fixed typography and dimensions. Below the comments, the line meter counts lines as NAVFIT98A prints them, for example 14 of 18 lines in NAVFIT98A. Comments allow 18 lines on EVAL and FITREP and 6 on CHIEFEVAL; blocks 28 and 29 allow 3 and 4; each CHIEFEVAL performance comment allows 2. The meter turns amber on the last line, red past the limit, and notes a last line that holds on" & _
            "e word. An Information finding lists each narrative past its limit; it never blocks saving, PDF output or NAVFIT98A export. Printed fit remains unverified. Check spelling, or F7 in a narrative block, runs Office spelling on this computer; accepted corrections become pending edits, and a signed report asks to reopen first. To stop Navy terms being flagged, save the list from Resources > Save Navy word list for spell check and add it under File > Options > Proofing > Custom Dictionaries; Swift never changes Office settings."
        Case 2: UXHelpText = "Reports provides search, report-type, summary-group, review and archive filters, saved views and selection actions. Click a table heading to sort ascending; click again to reverse. Open a report to inspect History, inspect workflow checkpoints and narrative saves. Routine administrative edits replace working state; narrative changes (main comments, employment/duties and Chief trait comments) remain in history. Routing, return, approval and signing retain checkpoints. Restore saves earlier values without erasing retained history. Duplicate creates a separate draft. Archived reports can be recovered; local approval must be reopened before editi" & _
            "ng."
        Case 3: UXHelpText = "Summary groups lists groups with report counts. Open a group for Members, Defaults, Review and History. Rename, archive and recover act on the group; reports remain independent. Defaults and shared fields use the official report blocks. Each changed field asks for confirmation and saves immediately. Defaults follow the report's Front Page layout for blocks 5-21 and 28; the group's reporting senior record supplies blocks 22-27 and 48, and block 29 is per Sailor. Group defaults replace that field in active members and future reports; global defaults also replace it in all active groups. Existing customized values are replaced. The confirmation " & _
            "identifies affected reports and local approvals returning to In-Work. Archived groups and reports are excluded. Cancel restores the field; failed saves roll back together. There is no separate preview or Apply step for defaults. Comparison uses editable trait cells and staged proposals; an internal baseline is captured automatically. Changed membership, revisions or locks require a fresh preview. Open All summary-group validation issues for member findings, grouping differences and provisional EP/MP limits with policy references. When creating reports from command members, filter by paygrade (such as E6) or specific rank/rate (such as YN1) in" & _
            " the single-report picker or batch creation screen. Changing a filter clears the prior member selection and invalidates a batch preview. The Closeout date column is the report period ending date: Mixed means different member ending dates, and Incomplete means one or more are missing. Member-table headings also sort rows; pending cells save before row order changes."
        Case 4: UXHelpText = "Resources > Reporting seniors and RSCA history stores reusable senior details. RSCA history is a table with E1-E9, W1-W5 and O1-O10 rows: enter RSCA, previous graded trait count (not reports), and the date verified/pulled from BOL. Leave unused rows blank. Save RSCA table records edited rows together; View prior baselines retains earlier observations. Editing the average or count uses a rounded-RSCA estimate rather than inventing an exact historical sum. Group pages show the reporting senior and provide Reporting senior / PSG RSCA plus Export group to NAVFIT98A. The calculator can open the selected senior record; choose the matching historica" & _
            "l baseline and confirm it excludes this group. A group projection includes subsequent local completions descended from its selected baseline, excluding its own previously counted contribution. Independent BOL observations require deliberate reconciliation. Apply senior details deliberately; directory edits do not silently update reports. Report > Report tools > Trait calculations separates pending and saved values. Arithmetic and manual provenance do not certify Navy policy compliance. Reporting senior, historical baseline and exclusion confirmation selections save automatically. Changing the senior or baseline clears the exclusion confirmati" & _
            "on. PSG RSCA recalculates from saved traits and appears beside Reporting senior on the group page. Apply reporting senior details to reports remains a separate deliberate action. Completed groups append local totals to the senior paygrade history. Counts are graded traits, excluding NOB. Re-completion applies only changes; reopening or archiving retains the recorded totals until corrected completion. Local totals are not BOL verification."
        Case 5: UXHelpText = "My Work lists assigned reviews, current ordered-route steps and ready local approvals. Its Report notes tab lists shared conversations separately from approval tasks. Each row names the member, report period, contributor, status and next action. Assignment due dates are shown when recorded; routes have no invented due dates. Future route steps stay out of the current contributor's queue. Requesters can see stale routes that need restarting. Select a row and choose Open selected work to reach that route or thread directly, or Open report to edit the report. Accept/Return apply only to an eligible selected assignment; stale and locked decisions" & _
            " are disabled. Completed work and all contributors are available through the history filter. Draft to complete: save the report, check validation, choose an ordered route in Review setup, then Start review. Each assigned contributor accepts a step or returns it with corrections. After changes, save and start a fresh route (restart a stale route as requester). After all steps are accepted, the final reviewer records Local approval, locking edits. Reopen approval with a reason to correct it. Complete the group separately after closeout. Report tools embeds Notes and replies, Review steps, Validation checks and Trait calculations. Notes and repl" & _
            "ies are independent of approval and do not require resolution or reopening. Unposted drafts remain private and can be resumed. Validation issues also opens as a modeless popout beside the report. Official signatures and submission remain separate."
        Case 6: UXHelpText = "The report action-area validation button and Report > Report tools > Validation checks check pending values live. Summary group > All summary-group validation issues includes all members and pending edits from the open member. Select a finding for its instruction, enclosure, printed page, paragraph/table and excerpt, then open the affected report block. Invalid entries, incomplete drafts, policy issues and Needs review are distinct. Storage checks are application rules. Group capacities are provisional until eligible membership is verified; no manual roster freeze is required. Live checks create no saved revisions or validation runs. Coverage" & _
            " is limited, including manual review of LDO/CWO categories, UIC exceptions and earlier/future instruction applicability. Reports > Draft checklist opens the report Validation checks view. Resources provides attributed narrative snippets. Reports > Import reports offers NAVFIT report or structured PDF import: Choose, Preview, Apply, Results. PDF mapping covers 41 EVAL, 52 CHIEFEVAL and 37 FITREP saved fields. Leave the merge target blank to create a new draft, or choose an existing draft and Fill empty fields / Replace mapped values. Full before/after previews show replacements and retained values. Blank or missing PDF fields never clear exist" & _
            "ing values. Close the target editor and resolve pending recovery before merging. Changed or locked targets require a fresh preview. To undo an import while no later edit exists, open the report's History and restore its Before import checkpoint, or archive a draft the import created. Use Check PDF extraction on this computer after choosing a PDF; extraction is local and does not require Acrobat or internet access. Changed sources or mapping choices require a fresh preview. Results can open the imported or merged report. Source field inspection and existing-draft reconciliation remain available. PDF imports use SSN by default; legacy DoDID map" & _
            "ping requires explicit compatibility selection. All imports retain provenance and are limited in scope. NAVFIT import maps 75 to 96 saved report values, depending on the report type, and retains all 128 native report columns, folders and summary values. Two-digit import years use 20xx automatically; no century selection is required. Select a summary group and choose Export group to NAVFIT98A. Choose an export directory; Swift supplies the native template and exports every saved group member. Incomplete content, findings and report status do not block export. Pending edits remain separate. Imported native-only values are preserved. Exact PSG R" & _
            "SCA is written when supported by complete traits and an exact historical sum; otherwise native values remain. Keep accompanying Swift baseline files locally. Open the native .accdb in NAVFIT98A for validation and printing. Returned-file reconciliation is not presented by this export screen. Counseling accepts a real ISO date, NOT REQ, or NOT PERF; NOT PERF requires an explanation. Live checks include counseling period, required career recommendation, partial-NOB trait limits, adverse-mark justification reminders and Regular reports exceeding 15 months (approval must be verified). Resources > NavyWriter catalog and Narrative snippets both filt" & _
            "er by rate. Legacy untagged snippets remain Unassigned; assign a rate when editing. Collection links open the complete source; only reviewed short excerpts can be imported. Shared readiness r110 uses the same implemented report checks in live findings, the saved draft checklist and NAVFIT export. Report tools > Validation checks and the group validation action open the full findings with references. New checks cover conditional counseling/PFA, duties, retention, senior details, rater relationships, periodic schedules (including the 2027 change), and known Regular continuity. Native-only/unknown catalogue cases remain Needs review; draft savin" & _
            "g stays available. Older saved runs are preserved under their historical pack."
        Case 7: UXHelpText = "Overview or Settings & Help > Recover pending edits opens checkpoints, separate from saved revisions. If the saved baseline changed, recover as a new draft. Settings & Help has backup/restore and a read-only operation history. Undo a saved change, including an import, from the report's History. To transfer all data from an older single-file Swift release, close the source, open the new release, and choose Import data from a previous Swift NAVFIT version on the setup screen (it makes the new data file first) or in Settings & Help. Select its .accdb, inspect the preview and choose Import all workspace data. After success, close and reopen the n" & _
            "ew app. Report values, IDs/relationships, revisions, local reviews, notes, defaults, members, prior evaluations, reporting seniors, RSCA, recovery and import provenance transfer together. The source is read-only; its forms/code are not copied. Existing populated destinations, future schemas and unknown source tables are rejected. Initial setup in an otherwise blank target is replaced. This is a workspace transfer, not a merge of separate workspaces. The command-line upgrade-workspace.ps1 uses the same migration engine. To update, open the new release; if it does not find your data file, choose Open an existing data file: a newer version upgra" & _
            "des an older data file automatically after copying it beside itself (""before schema N""), and puts it back from that copy if anything fails. A newer data file, or one open in another session, is refused unchanged. Back up the data file; the backup is checked row by row the same way. Settings & Help shows the workspace file size against the 2 GB Access limit and warns from 1.5 GB."
        Case 8: UXHelpText = "This is version {APP_VERSION}, schema {SCHEMA_VERSION}, a native Access development candidate. Automated native checks do not establish screen-reader or physical display/DPI acceptance. Desktop Access is verified within the documented tests; Free Access Runtime is unsupported; x86 remains unverified. No network integration, synchronization, authentication or official signatures are added. Settings & Help > Support and limitations provides technical diagnostics. Read CURRENT_STATE.md for exact artifact evidence and outstanding work."
        Case Else: UXHelpText = "Choose a topic."
    End Select
    UXHelpText = Replace(Replace(UXHelpText, "{APP_VERSION}", APP_VERSION), "{SCHEMA_VERSION}", CStr(SCHEMA_VERSION))
End Function
Public Function UXLoadHelp() As Boolean
    Dim i As Long
    Forms!frmUXHelp!lstTopics.RowSourceType = "Value List"
    Forms!frmUXHelp!lstTopics.RowSource = ""
    Forms!frmUXHelp!lstTopics.AddItem "0;Start here"
    Forms!frmUXHelp!lstTopics.AddItem "1;Create and save reports"
    Forms!frmUXHelp!lstTopics.AddItem "2;Find reports and history"
    Forms!frmUXHelp!lstTopics.AddItem "3;Summary groups and comparisons"
    Forms!frmUXHelp!lstTopics.AddItem "4;Reporting seniors and averages"
    Forms!frmUXHelp!lstTopics.AddItem "5;My Work and local review"
    Forms!frmUXHelp!lstTopics.AddItem "6;Checks and imports"
    Forms!frmUXHelp!lstTopics.AddItem "7;Recovery and backups"
    Forms!frmUXHelp!lstTopics.AddItem "8;Preview limits and diagnostics"
    Forms!frmUXHelp!lstTopics = 0
    Call UXHelpSelected
    UXLoadHelp = True
End Function
Public Function UXHelpSelected() As Boolean
    Forms!frmUXHelp!txtGuide = UXHelpText(Nz(Forms!frmUXHelp!lstTopics, -1))
    UXHelpSelected = True
End Function

' Context help never navigates away, flushes buffers or commits work.
Public Function UXWindowHelp(ByVal formName As String) As Boolean
    Dim topic As Long
    topic = UXWindowHelpTopic(formName)
    DoCmd.OpenForm "frmUXHelp"
    Forms!frmUXHelp!lstTopics = topic
    UXHelpSelected
    Forms!frmUXHelp!txtGuide = UXWindowHelpIntro(formName) & vbCrLf & vbCrLf & UXHelpText(topic)
    Forms!frmUXHelp!lblGuide.Caption = "Help for " & Forms(formName).Caption
    Forms!frmUXHelp.SetFocus
    UXWindowHelp = True
End Function
Public Function UXWindowHelpTopic(ByVal formName As String) As Long
    Select Case formName
        Case "frmEditor", "frmReportingSenior", "frmMemberDetails", "frmServiceDetails", "frmDecisionDetails", "frmReportPeople", "frmPartyNames", "frmRevisionCompare", "frmUXSignoff": UXWindowHelpTopic = 1
        Case "frmLibrary", "frmBulkFolders", "frmProvenance": UXWindowHelpTopic = 2
        Case "frmUXGroups", "frmMajorWorkbench", "frmUXBulk", "frmUXComparison", "frmUXSharedFields", "frmUXNewGroup", "frmUXMoveGroup", "frmUXRenameGroup", "frmUXGroupPolicy", "frmUXMembers", "frmUXGroupReport", "frmUXNewReport": UXWindowHelpTopic = 3
        Case "frmUXSeniors", "frmUXSeniorEditor", "frmUXGroupRSCA", "frmUXSeniorHistory", "frmUXSeniorPicker", "frmUXISICs": UXWindowHelpTopic = 4
        Case "frmMyWork", "frmUXReview", "frmReviewComments", "frmUXRouteTemplates", "frmUXContributors": UXWindowHelpTopic = 5
        Case "frmRecovery", "frmBackup", "frmOperations", "frmUXWorkspaceImport", "frmUXSettings": UXWindowHelpTopic = 7
        Case "frmSupport": UXWindowHelpTopic = 8
        Case "frmHome", "frmSetup": UXWindowHelpTopic = 0
        Case Else: UXWindowHelpTopic = 6
    End Select
End Function
Public Function UXWindowHelpLayout(ByVal formName As String, Optional ByVal viewportWidth As Long = 0) As Boolean
    Dim f As Access.Form, width As Long, c As Access.Control
    Set f = Forms(formName)
    If Not UXHasControl(f, "cmdWindowHelp") Then Exit Function
    'Match the width the workspace was just laid out for; a window resize passes none.
    width = viewportWidth
    If width = 0 Then width = f.InsideWidth
    If width < 1500 Then width = f.Width
    If width > UX_MAX_FORM_WIDTH Then width = UX_MAX_FORM_WIDTH
    UXVisualSet f!cmdWindowHelp, "Left", width - 800: UXVisualSet f!cmdWindowHelp, "Top", 150
    UXVisualSet f!cmdWindowHelp, "Width", 500: UXVisualSet f!cmdWindowHelp, "Height", 400
    If UXHasControl(f, "uxContext") Then
        'The editor's output actions share the context line's row.
        If UXHasControl(f, "uxPreviousReport") Then
            UXVisualSet f!uxContext, "Width", f!uxPreviousReport.Left - f!uxContext.Left - 150
        ElseIf UXHasControl(f, "uxCreatePDF") Then
            UXVisualSet f!uxContext, "Width", f!uxCreatePDF.Left - f!uxContext.Left - 150
        Else
            UXVisualSet f!uxContext, "Width", f!cmdWindowHelp.Left - f!uxContext.Left - 150
        End If
    End If
    For Each c In f.Controls
        If c.ControlType = acLabel And c.Top < 550 And c.Left < f!cmdWindowHelp.Left And c.Left + c.Width > f!cmdWindowHelp.Left Then
            UXVisualSet c, "Width", f!cmdWindowHelp.Left - c.Left - 150
        End If
    Next
    UXWindowHelpLayout = True
End Function
Public Function UXHelpResize(ByVal formName As String, ByVal original As String) As Boolean
    If Len(original) > 0 Then Eval original
    UXHelpResize = UXWindowHelpLayout(formName)
    UXApplyVisual Forms(formName)
End Function

Public Function UXWindowHelpIntro(ByVal formName As String) As String
    Select Case formName
        Case "frmUXDirectoryMerge"
            UXWindowHelpIntro = "Duplicate Sailor review: choose an original profile and a survivor. Surname matches only suggest candidates; choose any active survivor for a changed name. Preview masked identities and all report/prior links, independently confirm the same Sailor, and enter a reason without identifiers. Conflicting SSNs or DoD IDs block merging. Keep the SSN-owning profile as survivor. Apply moves stable member links, archives and retains the original profile, and preserves report fields, revisions, notes, review history and prior traits. Profile fields are not combined. Select merge history to undo only if no later affected changes exist. Close editors before merging or undoing. Folder permissions remain the storage protection; masking does not encrypt stored identifiers."
        Case "frmUXDirectory", "frmUXMemberEditor", "frmUXDirectoryBatch", "frmUXRUADImport", "frmUXConfirmIdentity"
            UXWindowHelpIntro = "Command directory: maintain member identities before creating reports. Preview batch creation or roster imports before applying. Changing a filter clears the earlier selection. Identity corrections do not silently replace saved report revisions."
        Case "frmUXMemberHistory", "frmUXPriorEditor", "frmUXCounselingEditor"
            UXWindowHelpIntro = "Member history: inspect prior evaluations separately from the current report. Verify dates, grade and provenance before saving historical entries. Counseling records follow the chapter 18 schedule; a recorded session is offered for blocks 30-31 of the member's next Regular report and is never saved into a report without your confirmation."
        Case "frmUXBulk"
            UXWindowHelpIntro = "Shared fields and defaults: read the scope banner before editing. Each field change has its own confirmation and immediate save; Cancel restores its previous value. Changed local approvals reopen only after confirmation."
        Case "frmUXComparison"
            UXWindowHelpIntro = "Members: edit trait cells in place. Queued cells save automatically; failed cells remain pending for retry or explicit discard. Resolve pending cells before leaving or changing groups."
        Case "frmUXNoteComposer", "frmUXNoteRecovery"
            UXWindowHelpIntro = "Draft notes: these are local notes, separate from the official narrative. Keep pending text until it is saved or deliberately discarded. Recovery checkpoints are not saved report revisions."
        Case "frmUXNavyWriter", "frmSnippets"
            UXWindowHelpIntro = "Narrative resources: filter by rate and inspect the source and attribution before inserting an excerpt. Review wording for this member; source collections are not automatic policy approval."
        Case "frmNAVFITExchange"
            UXWindowHelpIntro = "Export this summary group: choose an output directory and export all saved members. The internal native template is automatic. Findings do not block export; pending edits remain separate."
        Case "frmUXImportChoose", "frmPDFImport", "frmImportReview", "frmReconcile", "frmUXNativeGroups"
            UXWindowHelpIntro = "Import and exchange: choose the source, inspect the before/after preview, then apply deliberately. Changed sources or saved targets require a fresh preview. Keep the export baseline for return reconciliation."
        Case "frmUXReview", "frmReviewComments"
            UXWindowHelpIntro = "Review: decisions use saved values and require the assigned contributor. Enter a decision note. Stale routes need the requester to restart them. Local approval is separate from group status and official submission."
        Case "frmUXSignoff"
            UXWindowHelpIntro = "Mark signed: record the date the official report was signed outside Swift. The saved version becomes the signed copy and the report turns read only; editing a field later asks whether to reopen it for correction. Disposing of drafts removes superseded draft versions, recovery checkpoints, unposted note drafts and imported source values; the signed version and every conversation stay."
        Case "frmEditor"
            UXWindowHelpIntro = "Report editor: Front and Back preserve pending edits. Save draft commits a revision; recovery only checkpoints unsaved work. Opening or closing this guide leaves your report and pending values in place."
        Case "frmUXSettings"
            UXWindowHelpIntro = "Settings and help: open the complete workspace guide here, or use the top-right ? in any window for its workflow. Recovery, backups and previous-version transfer remain separate guarded tools."
        Case Else
            UXWindowHelpIntro = "This guide explains the workflow for this window. Choose another topic on the left for related tasks. Close guide to return to your unchanged work."
    End Select
End Function
