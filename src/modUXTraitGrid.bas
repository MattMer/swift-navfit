Attribute VB_Name = "modUXTraitGrid"
Option Explicit
Private cacheWorkspace As DAO.Workspace
Private cacheDB As DAO.Database
Private cacheRows As DAO.Recordset
Private cachePath As String
Private loading As Boolean
Private saving As Boolean
Private queuedCells As Object
Private liveGroup As Long
Private rowStates As Object
' Return the number of unsaved report/field cells, not the number of saved revisions.
Public Function UXGridPendingCount() As Long
    If Not queuedCells Is Nothing Then UXGridPendingCount = queuedCells.Count
End Function
' Show retry/discard only while pending cells exist; move focus before hiding a focused button.
Private Sub UXGridPendingActions()
    Dim activeName As String
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Sub
    If UXGridPendingCount() = 0 Then
        On Error Resume Next
        activeName = Forms!frmUXComparison.ActiveControl.Name
        On Error GoTo 0
        If activeName = "cmdRetryCells" Or activeName = "cmdDiscardCells" Then Forms!frmUXComparison!cmdBack.SetFocus
    End If
    With Forms!frmUXComparison
        If UXHasControl(Forms!frmUXComparison, "cmdRetryCells") Then .cmdRetryCells.Visible = (UXGridPendingCount() > 0)
        If UXHasControl(Forms!frmUXComparison, "cmdDiscardCells") Then .cmdDiscardCells.Visible = (UXGridPendingCount() > 0)
    End With
End Sub
' Release the separate cache workspace and temporary file only when no pending cells remain.
Public Function UXGridClose() As Boolean
    If UXGridPendingCount() > 0 Then Exit Function
    loading = True
    On Error Resume Next
    If Not cacheRows Is Nothing Then cacheRows.Close
    Set cacheRows = Nothing
    If Not cacheDB Is Nothing Then cacheDB.Close
    Set cacheDB = Nothing
    If Not cacheWorkspace Is Nothing Then cacheWorkspace.Close
    Set cacheWorkspace = Nothing
    saving = False: Set queuedCells = Nothing
    If Len(cachePath) > 0 Then Kill cachePath
    If Len(cachePath) > 0 Then
        If Len(Dir$(cachePath)) = 0 Then cachePath = ""
    End If
    UXGridClose = True
End Function
' Build the editable cache from live group reports or a historical comparison baseline.
' Use a separate DAO workspace to avoid locking the report save transaction; cached row
' state serves display-time guards. Refuse reload while unsaved queue values exist.
Public Function UXGridLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, p As DAO.Recordset, db As DAO.Database, sql As String, block As Long, id As Long, kind As String, previous As Object
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, , "Retry or discard pending cells before refreshing the comparison."
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    loading = True
    Set f.Recordset = Nothing
    UXGridClose
    cachePath = Environ$("TEMP") & "\Swift-NAVFIT-traits-" & Replace(NewIdentity(), "-", "") & ".accdb"
    Set cacheWorkspace = DBEngine.CreateWorkspace("SwiftComparisonCache", "admin", "", dbUseJet)
    Set cacheDB = cacheWorkspace.CreateDatabase(cachePath, dbLangGeneral, dbVersion120)
    sql = "CREATE TABLE PendingTraits (ReportID LONG PRIMARY KEY, SortOrdinal LONG, RevisionID LONG, Member TEXT(120), RowStatus TEXT(120), ObservationStatus TEXT(8), ITA TEXT(12)"
    For block = 33 To 39: sql = sql & ", Trait" & block & " TEXT(3)": Next
    sql = sql & ", PromotionRecommendation TEXT(20), AssignmentRecommendationA TEXT(255), AssignmentRecommendationB TEXT(255), PriorITA TEXT(12), PriorPromotion TEXT(30), PriorPeriod TEXT(32), PriorSource TEXT(255))": cacheDB.Execute sql, dbFailOnError
    Set cacheRows = cacheDB.OpenRecordset("PendingTraits", dbOpenDynaset)
    Set rowStates = CreateObject("Scripting.Dictionary")
    Set db = CurrentDb
    liveGroup = 0
    If Left$(Nz(Forms!frmUXComparison.OpenArgs, ""), 2) = "G:" And Nz(Forms!frmUXComparison!txtMeeting, 0) = 0 Then
        liveGroup = CLng(Mid$(Forms!frmUXComparison.OpenArgs, 3))
        Set r = db.OpenRecordset("SELECT r.*,r.CurrentRevisionID AS RevisionID,r.ReportID AS SortOrdinal FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & liveGroup & " ORDER BY r.ReportID", dbOpenSnapshot)
    Else
        Set r = db.OpenRecordset("SELECT m.*,r.IsArchived,r.WorkflowState,r.CurrentRevisionID FROM ComparisonMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.SetID=" & Forms!frmUXComparison!txtMeeting & " ORDER BY m.SortOrdinal", dbOpenSnapshot)
    End If
    Do While Not r.EOF
        cacheRows.AddNew: id = r!ReportID
        rowStates(CStr(id)) = Array(CBool(r!IsArchived), CStr(r!WorkflowState), CLng(Nz(r!CurrentRevisionID, 0)))
        cacheRows!ReportID = id: cacheRows!SortOrdinal = r!SortOrdinal: cacheRows!RevisionID = r!RevisionID
        cacheRows!Member = UXRevisionMemberLabel(r!RevisionID)
        If liveGroup > 0 Then
            cacheRows!ObservationStatus = r!ObservationStatus
            For block = 33 To 39: cacheRows.Fields("Trait" & block) = r.Fields("Trait" & block): Next
            cacheRows!PromotionRecommendation = r!PromotionRecommendation
            cacheRows!AssignmentRecommendationA = r!AssignmentRecommendationA
            cacheRows!AssignmentRecommendationB = r!AssignmentRecommendationB
        Else
            cacheRows!ObservationStatus = MajorRevisionValue(r!RevisionID, "ObservationStatus")
            For block = 33 To 39: cacheRows.Fields("Trait" & block) = MajorRevisionValue(r!RevisionID, "Trait" & block): Next
            cacheRows!PromotionRecommendation = MajorRevisionValue(r!RevisionID, "PromotionRecommendation")
            cacheRows!AssignmentRecommendationA = MajorRevisionValue(r!RevisionID, "AssignmentRecommendationA")
            cacheRows!AssignmentRecommendationB = MajorRevisionValue(r!RevisionID, "AssignmentRecommendationB")
        End If
        Set previous = PreviousMemberEvaluation(id)
        cacheRows!PriorITA = previous("ITA"): cacheRows!PriorPromotion = previous("Promotion"): cacheRows!PriorPeriod = previous("Period"): cacheRows!PriorSource = previous("Source")
        Set p = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " AND ReportID=" & id & " AND ProposalState='Proposed' ORDER BY ProposalID", dbOpenSnapshot)
        Do While Not p.EOF
            If p!IsNullValue Then cacheRows.Fields(p!FieldKey) = Null Else cacheRows.Fields(p!FieldKey) = p!ValueText
            p.MoveNext
        Loop
        p.Close
        cacheRows!RowStatus = UXGridLockReason(id, r!RevisionID)
        If Len(cacheRows!RowStatus) = 0 Then cacheRows!RowStatus = "Editable - saves automatically"
        cacheRows.Update: r.MoveNext
    Loop
    r.Close
    cacheRows.Close
    Set cacheRows = cacheDB.OpenRecordset("SELECT * FROM PendingTraits ORDER BY SortOrdinal", dbOpenDynaset)
    If Not cacheRows.EOF Then cacheRows.MoveLast: cacheRows.MoveFirst
    Set f.Recordset = cacheRows
    kind = Nz(DLookup("FormKind", "Report", "ReportID=" & Nz(f!ReportID, 0)), "EVAL")
    For block = 33 To 39
        f.Controls("heading" & block).Caption = UXFieldCaption("Trait" & block, kind)
        f.Controls("Trait" & block).ControlTipText = UXFieldCaption("Trait" & block, kind)
    Next
    f!headingAssignmentRecommendationA.Caption = UXFieldCaption("AssignmentRecommendationA", kind)
    f!headingAssignmentRecommendationB.Caption = UXFieldCaption("AssignmentRecommendationB", kind)
    Call UXGridRefreshAverages: loading = False
    Call UXGridCurrent: Call UXGridViewport: UXGridLoad = True
End Function
' Use loaded archive/workflow/revision state for display without database reads. With
' authoritative=True, recheck the report and live membership or historical group stamp
' at the write boundary; return an empty string when editing is allowed.
Public Function UXGridLockReason(ByVal id As Long, ByVal revision As Long, Optional ByVal authoritative As Boolean = False) As String
    Dim snapshot As Long, groupID As Long, state As Variant, db As DAO.Database, r As DAO.Recordset
    If authoritative Then
        Set db = CurrentDb
        Set r = db.OpenRecordset("SELECT CurrentRevisionID,IsArchived,WorkflowState FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
        If r.EOF Then r.Close: UXGridLockReason = "Report removed - reopen members": Exit Function
        state = Array(CBool(r!IsArchived), CStr(r!WorkflowState))
        If Nz(r!CurrentRevisionID, 0) <> revision Then r.Close: UXGridLockReason = "Saved report changed - reopen comparison": Exit Function
        r.Close
        If liveGroup > 0 Then
            If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & liveGroup), True) Or DCount("*", "GroupMember", "GroupID=" & liveGroup & " AND ReportID=" & id) = 0 Then UXGridLockReason = "Group membership changed - reopen members": Exit Function
        Else
            snapshot = Nz(DLookup("GroupSnapshotID", "ComparisonSet", "SetID=" & Forms!frmUXComparison!txtMeeting), 0)
            groupID = Nz(DLookup("GroupID", "GroupSnapshot", "SnapshotID=" & snapshot), 0)
            If Nz(DLookup("Stamp", "GroupSnapshot", "SnapshotID=" & snapshot), "") <> GroupStamp(groupID) Then UXGridLockReason = "Group changed - reopen comparison": Exit Function
        End If
    Else
        If rowStates Is Nothing Then Exit Function
        If Not rowStates.Exists(CStr(id)) Then Exit Function
        state = rowStates(CStr(id))
        If state(2) <> revision Then UXGridLockReason = "Saved report changed - reopen comparison": Exit Function
    End If
    If state(0) Then UXGridLockReason = "Archived - recover report to edit": Exit Function
    If state(1) <> "Draft" And state(1) <> "Approved" Then UXGridLockReason = "Review lock - reopen report"
End Function
' Apply cached row locks and contextual status on row focus; avoid whole-group queries.
Public Function UXGridCurrent() As Boolean
    Dim f As Access.Form, reason As String, block As Long
    If loading Or saving Or cacheRows Is Nothing Then Exit Function
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If f.Recordset.EOF Then Exit Function
    f!PriorITA.ControlTipText = Nz(f!PriorPeriod, "") & " - " & Nz(f!PriorSource, "")
    f!PriorPromotion.ControlTipText = f!PriorITA.ControlTipText
    reason = UXGridLockReason(f!ReportID, f!RevisionID)
    For block = 33 To 39: f.Controls("Trait" & block).Locked = (Len(reason) > 0): Next
    f!PromotionRecommendation.Locked = (Len(reason) > 0)
    f!AssignmentRecommendationA.Locked = (Len(reason) > 0)
    f!AssignmentRecommendationB.Locked = (Len(reason) > 0)
    If UXGridPendingCount() = 0 Then Forms!frmUXComparison!lblStatus.Caption = IIf(Len(reason) > 0, reason, "Edit cells directly. Valid changes save automatically as report revisions.")
    UXGridCurrent = True
End Function
' Coalesce each report/field edit and capture its contributor. A short timer drains
' the queue after the bound control event finishes, avoiding nested save/lock problems.
Public Function UXGridQueue(ByVal key As String) As Boolean
    Dim f As Access.Form, identity As String
    On Error GoTo Failed
    If loading Or saving Then Exit Function
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If queuedCells Is Nothing Then Set queuedCells = CreateObject("Scripting.Dictionary")
    identity = CStr(f!ReportID) & ":" & key
    queuedCells(identity) = Array(CLng(f!ReportID), key, f.Controls(key).Value, MajorAuthor())
    Forms!frmUXComparison.TimerInterval = 150
    Forms!frmUXComparison!lblStatus.Caption = "Saving pending cell changes..."
    UXGridPendingActions
    UXGridQueue = True: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = "Queue cell save: " & Err.Number & " " & Err.Description
End Function
' Save queued cells in order and remove each only after success. Stop at the first
' failure, retaining that cell and later values; already committed revisions remain saved.
Public Function UXGridDrain() As Boolean
    Dim pending As Object, item As Variant, cell As Variant, f As Access.Form, r As DAO.Recordset, ok As Boolean, failure As String
    On Error GoTo Failed
    If saving Then Exit Function
    If UXGridPendingCount() = 0 Then UXGridDrain = True: Exit Function
    Forms!frmUXComparison.TimerInterval = 0
    Set pending = queuedCells
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    'Finish Access's bound-cache update before entering any report transaction.
    If f.Dirty Then f.Dirty = False
    ok = True
    For Each item In pending.Keys
        cell = pending(item)
        If CLng(cell(3)) <> MajorAuthor() Then Err.Raise vbObjectError + 4815, , "Pending cells belong to another contributor. Switch back before retrying or discard these cells."
        Set r = f.RecordsetClone: r.FindFirst "ReportID=" & cell(0)
        If r.NoMatch Then Err.Raise vbObjectError + 4813, , "Pending report is no longer in the comparison."
        f.Bookmark = r.Bookmark: r.Close
        f.Controls(CStr(cell(1))) = cell(2)
        If Not UXGridEdit(CStr(cell(1))) Then ok = False: Exit For
        pending.Remove CStr(item)
    Next
    If Not ok Then
        failure = Forms!frmUXComparison!lblStatus.Caption
        UXGridRetainPending failure
    Else
        Set queuedCells = Nothing
        UXGridPendingActions
    End If
    UXGridDrain = ok: Exit Function
Failed:
    failure = "Pending cell save: " & Err.Number & " " & Err.Description
    UXGridRetainPending failure
End Function
' Restore intended unsaved values to the cache after a failed drain and stop automatic
' retries until the user chooses Retry or Discard.
Private Sub UXGridRetainPending(ByVal failure As String)
    Dim r As DAO.Recordset, key As Variant, cell As Variant, restoreFailure As String
    On Error GoTo Failed
    Forms!frmUXComparison.TimerInterval = 0
    Set r = cacheDB.OpenRecordset("PendingTraits", dbOpenDynaset)
    For Each key In queuedCells.Keys
        cell = queuedCells(key): r.FindFirst "ReportID=" & cell(0)
        If Not r.NoMatch Then
            r.Edit: r.Fields(CStr(cell(1))) = cell(2): r!RowStatus = "Not saved - retry or discard pending cells": r.Update
        End If
    Next
    r.Close: Set r = Nothing
    Forms!frmUXComparison!uxTraitGrid.Form.Refresh
Finished:
    Forms!frmUXComparison!lblStatus.Caption = failure & " " & UXGridPendingCount() & " cells remain unsaved; earlier successful cells are saved. Use Retry pending cells or Discard pending cells." & restoreFailure
    UXGridPendingActions: Exit Sub
Failed:
    restoreFailure = " Pending values remain in memory; display refresh failed: " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    On Error GoTo 0
    GoTo Finished
End Sub
' Discard only the unsaved queue, reload saved values, and requery native rows so the
' displayed cell values reset too; committed report revisions are preserved.
Public Function UXGridDiscardPending() As Boolean
    On Error GoTo Failed
    Forms!frmUXComparison.TimerInterval = 0
    If Forms!frmUXComparison!uxTraitGrid.Form.Dirty Then Forms!frmUXComparison!uxTraitGrid.Form.Undo
    Set queuedCells = Nothing
    If Not UXGridLoad() Then Exit Function
    Forms!frmUXComparison!uxTraitGrid.Form.Requery
    UXGridPendingActions
    Forms!frmUXComparison!lblStatus.Caption = "Unsaved cells discarded. Previously saved report revisions remain."
    UXGridDiscardPending = True: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = "Discard display refresh failed: " & Err.Number & " " & Err.Description & ". Close and reopen the comparison."
End Function
' Validate and atomically save one cell against its expected revision, optionally
' reopening local approval. Live Members updates its cache without copying group history;
' explicit historical comparisons retain baseline refreshes. Roll back failed writes.
Public Function UXGridEdit(ByVal key As String, Optional ByVal reopenDecision As Long = 0) As Boolean
    Dim f As Access.Form, value As Variant, reason As String, w As DAO.Workspace, started As Boolean, db As DAO.Database, before As Variant, proposal As Long, failure As String, committed As Boolean, stage As String, expectedRevision As Long, baselineRead As Boolean
    If loading Or saving Then Exit Function
    On Error GoTo Failed
    stage = "Read edited cell"
    Set f = Forms!frmUXComparison!uxTraitGrid.Form: value = f.Controls(key).Value
    If Not IsTraitKey(key) And key <> "PromotionRecommendation" And key <> "AssignmentRecommendationA" And key <> "AssignmentRecommendationB" Then Err.Raise vbObjectError + 4812, , "Only traits, promotion and career recommendations can be edited here."
    If key = "AssignmentRecommendationA" Or key = "AssignmentRecommendationB" Then
        If Len(Nz(value, "")) > 20 Then Err.Raise vbObjectError + 4811, , "Career recommendations allow up to 20 characters."
    End If
    stage = "Check edit guards"
    reason = UXGridLockReason(f!ReportID, f!RevisionID, True)
    If key = "PromotionRecommendation" Then
        Dim validated As Variant
        validated = ValidateAssessment(Array(f!ObservationStatus.Value, value))
    End If
    If Len(reason) > 0 Then Err.Raise vbObjectError + 4810, , reason
    If Left$(key, 5) = "Trait" And Not IsNull(value) Then
        If Len(value) > 0 And InStr(1, ";NOB;1.0;2.0;3.0;4.0;5.0;", ";" & value & ";", vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 4811, , "Choose NOB or a whole mark from 1.0 to 5.0."
    End If
    stage = "Read saved baseline"
    expectedRevision = CLng(f!RevisionID)
    before = MajorRevisionValue(expectedRevision, key): baselineRead = True
    If SameFieldValue(before, value) Then UXGridEdit = True: Exit Function
    Dim reportID As Long, groupID As Long, snapshot As Long, newSet As Long, revision As Long, changes As Object
    reportID = f!ReportID
    If liveGroup > 0 Then
        groupID = liveGroup
    Else
        snapshot = Nz(DLookup("GroupSnapshotID", "ComparisonSet", "SetID=" & Forms!frmUXComparison!txtMeeting), 0)
        groupID = Nz(DLookup("GroupID", "GroupSnapshot", "SnapshotID=" & snapshot), 0)
    End If
    saving = True
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "") = "Approved" Then
        If reopenDecision = 0 Then reopenDecision = MsgBox("This report has a local approval. Set it back to In-Work and save this change? The approval will remain in history.", vbYesNo + vbQuestion, "Reopen approved report")
        If reopenDecision <> vbYes Then
            f.Undo: saving = False: UXGridEdit = False: Exit Function
        End If
    End If
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add key, value
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    stage = "Finish comparison cache edit"
    If f.Dirty Then f.Dirty = False
    stage = "Begin save transaction"
    w.BeginTrans: started = True
    stage = "Reopen local approval"
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "") = "Approved" Then ReopenApprovedReport reportID, "Contributor changed " & key & " in group members"
    stage = "Save report revision"
    Call CommitSavedFields(reportID, expectedRevision, changes, "Automatic group member cell save: " & key)
    revision = CurrentRevision(reportID)
    newSet = Nz(Forms!frmUXComparison!txtMeeting, 0)
    If liveGroup = 0 Then
        stage = "Refresh group baseline"
        snapshot = FreezeSummaryGroup(groupID, "Comparison refreshed after automatic cell save")
        stage = "Refresh comparison baseline"
        newSet = CreateComparisonSet(snapshot, "Group members - automatic save")
    End If
    stage = "Flush edited comparison row"
    If f.Dirty Then f.Dirty = False
    w.CommitTrans: started = False: committed = True
    stage = "Refresh saved row"
    Forms!frmUXComparison!txtMeeting = newSet
    rowStates(CStr(reportID)) = Array(False, "Draft", revision)
    f!RevisionID = revision: f!RowStatus = "Saved automatically": If f.Dirty Then f.Dirty = False
    stage = "Refresh averages"
    UXGridRefreshAverages
    stage = "Refresh comparison preview"
    UXRefreshComparisonPreview
    Forms!frmUXComparison!lblStatus.Caption = "Saved automatically at " & Format$(Now(), "h:nn:ss AM/PM")
    saving = False: UXGridEdit = True: Exit Function
Failed:
    failure = stage & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    If started Then w.Rollback
    saving = False
    If committed Then
        Forms!frmUXComparison!lblStatus.Caption = "Report saved, but the comparison refresh failed: " & failure & ". Close and reopen the comparison."
        UXGridEdit = True: Exit Function
    End If
    On Error Resume Next
    f.Undo
    If baselineRead Then
        f.Controls(key) = before: f!RevisionID = expectedRevision
        If f.Dirty Then f.Dirty = False
        UXGridRefreshAverages
    End If
    Forms!frmUXComparison!lblStatus.Caption = failure & " Cell change was not saved."
End Function
' Compute pending cache ITAs and the group trait average, excluding wholly NOB reports.
' Write an ITA cache row only when its formatted value changed.
Public Function UXGridRefreshAverages() As Boolean
    Dim r As DAO.Recordset, block As Long, count As Long, sum As Double, rowCount As Long, rowSum As Double, value As String, nob As Long
    Set r = cacheDB.OpenRecordset("PendingTraits", dbOpenDynaset)
    Do While Not r.EOF
        rowCount = 0: rowSum = 0
        If Nz(r!ObservationStatus, "") <> "NOB" Then
            For block = 33 To 39
                value = Nz(r.Fields("Trait" & block), "")
                If value <> "NOB" And Len(value) > 0 And IsNumeric(value) Then rowCount = rowCount + 1: rowSum = rowSum + CDbl(value)
            Next
        Else
            nob = nob + 1
        End If
        Dim average As String
        If rowCount > 0 Then average = Format$(rowSum / rowCount, "0.00") Else average = "-"
        If Nz(r!ITA, "") <> average Then
            r.Edit: r!ITA = average: r.Update
        End If
        sum = sum + rowSum: count = count + rowCount: r.MoveNext
    Loop
    r.Close
    Forms!frmUXComparison!uxGridAverage.Caption = "Group trait average: " & IIf(count > 0, Format$(sum / IIf(count = 0, 1, count), "0.00"), "Unavailable") & " | " & count & " numeric traits | " & nob & " whole-report NOB excluded"
    Forms!frmUXComparison!uxTraitGrid.Form.Refresh
    UXRefreshGroupSenior
    UXGridRefreshAverages = True
End Function
' Populate full before/after proposal rows for the current comparison; optionally include applied results.
Public Function UXRefreshComparisonPreview(Optional ByVal results As Boolean = False) As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXComparison
    f!lstProposals.RowSource = "SELECT p.ProposalID,UXIdentity(r.OfficialName,r.MemberName) AS Member,UXFieldCaption(p.FieldKey,r.FormKind) AS Block,MajorRevisionValue(p.BaseRevision,p.FieldKey) AS [Before],p.ValueText AS [After],p.ProposalState AS State FROM MeetingProposal AS p INNER JOIN Report AS r ON p.ReportID=r.ReportID WHERE p.SetID=" & f!txtMeeting & " AND p.ProposalState IN (" & IIf(results, "'Proposed','Applied'", "'Proposed'") & ") ORDER BY UXIdentity(r.OfficialName,r.MemberName),p.FieldKey"
    f!lstProposals.Requery: UXRefreshComparisonPreview = True
End Function
' Show the selected proposal's complete before/after values outside the condensed list.
Public Function UXGridDetail() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, db As DAO.Database
    Set f = Forms!frmUXComparison
    If Nz(f!lstProposals, 0) = 0 Then Exit Function
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE ProposalID=" & f!lstProposals, dbOpenSnapshot)
    f!uxProposalDetail = UXFieldCaption(r!FieldKey, Nz(DLookup("FormKind", "Report", "ReportID=" & r!ReportID), "EVAL")) & vbCrLf & "BEFORE: " & Nz(MajorRevisionValue(r!BaseRevision, r!FieldKey), "(blank)") & vbCrLf & "AFTER: " & Nz(r!ValueText, "(clear value)")
    r.Close: UXGridDetail = True
End Function

' Fit the native grid columns to the viewport and retain the live or historical group context.
Public Function UXGridViewport() As Boolean
    Dim f As Access.Form, c As Access.Control, v As Variant, usable As Long, ratio As Double
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    UXResizeWorkspace "frmUXComparison"
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    usable = Forms!frmUXComparison!uxTraitGrid.Width - 350: ratio = usable / 17050
    For Each c In f.Controls
        If InStr(c.Tag, ":") > 0 Then
            v = Split(c.Tag, ":"): c.Left = CLng(v(0) * ratio): c.Width = CLng(v(1) * ratio)
        End If
    Next
    f.Width = usable: f.Section(acHeader).Height = IIf(usable < 9000, 1500, 950)
    If liveGroup = 0 Then Forms!frmUXComparison!uxGroupContext = Nz(MajorScalar("SELECT g.GroupID FROM GroupSnapshot AS g INNER JOIN ComparisonSet AS c ON g.SnapshotID=c.GroupSnapshotID WHERE c.SetID=" & Nz(Forms!frmUXComparison!txtMeeting, 0)), 0)
    UXGridViewport = True
End Function

' Drain pending edits before detaching the cache; False tells Form_Unload to cancel the close.
Public Function UXGridUnload(ByVal f As Access.Form) As Boolean
    If Not UXGridDrain() Then Exit Function
    loading = True
    Set f!uxTraitGrid.Form.Recordset = Nothing
    UXGridClose
    UXGridUnload = True
End Function
' Expose the cache path for diagnostics and cleanup verification.
Public Function UXGridCacheFile() As String
    UXGridCacheFile = cachePath
End Function

' Close through the pending-cell unload guard. Normal live exits return to Summary groups;
' page transitions pass revealDestination=False to avoid exposing an intermediate window.
Public Function UXCloseComparison(Optional ByVal revealDestination As Boolean = True) As Boolean
    Dim isLive As Boolean
    On Error GoTo Failed
    isLive = Left$(Nz(Forms!frmUXComparison.OpenArgs, ""), 2) = "G:"
    DoCmd.Close acForm, "frmUXComparison", acSaveNo
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    If Not revealDestination Then UXCloseComparison = True: Exit Function
    If isLive Then
        OpenScreen "frmUXGroups"
    ElseIf CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then
        Forms!frmMajorWorkbench.Visible = True
        Forms!frmMajorWorkbench.SetFocus
    Else
        OpenScreen "frmMajorWorkbench"
    End If
    UXCloseComparison = True
    Exit Function
Failed:
    If Err.Number = 2501 Then Exit Function
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then Forms!frmUXComparison!lblStatus.Caption = "Close comparison: " & Err.Number & " " & Err.Description
End Function
