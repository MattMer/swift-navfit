Attribute VB_Name = "modUXTraitGrid"
Option Explicit
Private gridSortKey As String, gridSortDescending As Boolean

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

'Read a PDF snapshot without draining, discarding or changing the bound cache.
Public Function UXGridPDFSnapshot(ByVal reportID As Long, ByVal fields As Object) As Boolean
    Dim pending As Variant, key As Variant, names As Variant, i As Long, f As Access.Form, value As Variant, active As Access.Control
    If cacheDB Is Nothing Then Exit Function
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    names = Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39", "PromotionRecommendation", "AssignmentRecommendationA", "AssignmentRecommendationB")
    If Not queuedCells Is Nothing Then
        For Each key In queuedCells.Keys
            pending = queuedCells(key)
            If CLng(pending(0)) = reportID Then
                fields(CStr(pending(1))) = pending(2): UXGridPDFSnapshot = True
            End If
        Next
    End If
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If f.Recordset.EOF Then Exit Function
    If Nz(f!ReportID, 0) <> reportID Then Exit Function
    If Not f.Dirty Then Exit Function
    'Only the active unfinished cell supplements the queue; unchanged cached
    'columns must not overwrite another pending editor field.
    On Error Resume Next
    Set active = Screen.ActiveControl
    On Error GoTo 0
    If active Is Nothing Then Exit Function
    If Not active.Parent Is f Then Exit Function
    For i = 0 To UBound(names)
        If active.Name = names(i) Then
            value = active.Value
            On Error Resume Next
            value = active.Text
            On Error GoTo 0
            fields(CStr(names(i))) = value: UXGridPDFSnapshot = True
            Exit For
        End If
    Next
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
' The cache holds report values, so a file that survives is reported and its path kept
' for a later close to retry; closing the derived cursors stays best effort.
Public Function UXGridClose() As Boolean
    Dim failure As Object
    If UXGridPendingCount() > 0 Then Exit Function
    loading = True
    On Error Resume Next
    If Not cacheRows Is Nothing Then cacheRows.Close
    Set cacheRows = Nothing
    If Not cacheDB Is Nothing Then cacheDB.Close
    Set cacheDB = Nothing
    If Not cacheWorkspace Is Nothing Then cacheWorkspace.Close
    Set cacheWorkspace = Nothing
    Err.Clear
    saving = False: Set queuedCells = Nothing
    If Len(cachePath) > 0 Then
        Kill cachePath
        If Err.Number <> 0 Then Set failure = SwiftFailure("modUXTraitGrid", "UXGridClose", "Delete trait cache", Err.Number, Err.Description, Err.Source)
        Err.Clear
        If Len(Dir$(cachePath)) = 0 Then cachePath = ""
    End If
    On Error GoTo 0
    If Len(cachePath) > 0 Then
        If failure Is Nothing Then Set failure = SwiftFailure("modUXTraitGrid", "UXGridClose", "Delete trait cache", vbObjectError + 4816, "The comparison cache file could not be removed.", "modUXTraitGrid.UXGridClose")
        Exit Function
    End If
    UXGridClose = True
End Function
' Build the editable cache from live group reports or a historical comparison baseline.
' Use a separate DAO workspace to avoid locking the report save transaction; cached row
' state serves display-time guards. Refuse reload while unsaved queue values exist.
Public Function UXGridLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, p As DAO.Recordset, db As DAO.Database, sql As String, block As Long, id As Long, kind As String, previous As Object, priorRows As Object, savedValues As Object, proposals As Object, entries As Collection, entry As Variant, reportSQL As String, revisionSQL As String, trans As Boolean, loadFailure As Object, savedValue As Variant, memberLabel As String
    Call UXAssertView(UX_MEMBERS_VIEW): Call UXAssertView(UX_TRAIT_ROWS_VIEW)
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, "modUXTraitGrid.UXGridLoad", "Retry or discard pending cells before refreshing the comparison."
    On Error GoTo Failed
    gridSortKey = "": gridSortDescending = False
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
        reportSQL = "SELECT ReportID FROM GroupMember WHERE GroupID=" & liveGroup
        revisionSQL = "SELECT CurrentRevisionID FROM Report WHERE ReportID In (" & reportSQL & ")"
        Set r = db.OpenRecordset("SELECT r.*,r.CurrentRevisionID AS RevisionID,r.ReportID AS SortOrdinal FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & liveGroup & " ORDER BY r.ReportID", dbOpenSnapshot)
    Else
        reportSQL = "SELECT ReportID FROM ComparisonMember WHERE SetID=" & Forms!frmUXComparison!txtMeeting
        revisionSQL = "SELECT RevisionID FROM ComparisonMember WHERE SetID=" & Forms!frmUXComparison!txtMeeting
        Set r = db.OpenRecordset("SELECT m.*,r.IsArchived,r.WorkflowState,r.CurrentRevisionID FROM ComparisonMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE m.SetID=" & Forms!frmUXComparison!txtMeeting & " ORDER BY m.SortOrdinal", dbOpenSnapshot)
    End If
    Set priorRows = PreviousMemberEvaluationBatch(reportSQL)
    Set savedValues = CreateObject("Scripting.Dictionary"): Set proposals = CreateObject("Scripting.Dictionary")
    'Joins, not RevisionID In (subquery): the engine scans RevisionValue for the latter.
    If liveGroup > 0 Then
        sql = "SELECT v.RevisionID,v.FieldKey,v.ValueText,v.IsNullValue FROM (GroupMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID) INNER JOIN RevisionValue AS v ON r.CurrentRevisionID=v.RevisionID WHERE m.GroupID=" & liveGroup
    Else
        sql = "SELECT v.RevisionID,v.FieldKey,v.ValueText,v.IsNullValue FROM ComparisonMember AS c INNER JOIN RevisionValue AS v ON c.RevisionID=v.RevisionID WHERE c.SetID=" & CLng(Forms!frmUXComparison!txtMeeting)
    End If
    sql = sql & " AND v.FieldKey In ('OfficialName','MemberName'"
    If liveGroup = 0 Then sql = sql & ",'ObservationStatus','Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39','PromotionRecommendation','AssignmentRecommendationA','AssignmentRecommendationB'"
    Set p = db.OpenRecordset(sql & ")", dbOpenSnapshot)
    Do While Not p.EOF
        If p!IsNullValue Then savedValues(p!RevisionID & ":" & p!FieldKey) = Null Else savedValues(p!RevisionID & ":" & p!FieldKey) = p!ValueText.Value
        p.MoveNext
    Loop
    p.Close
    If liveGroup = 0 Then
        Set p = db.OpenRecordset("SELECT ReportID,FieldKey,ValueText,IsNullValue FROM qryMeetingProposalRevisionIDs WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " AND ProposalState='Proposed' ORDER BY ProposalID", dbOpenSnapshot)
        Do While Not p.EOF
            If Not proposals.Exists(CStr(p!ReportID)) Then Set entries = New Collection: proposals.Add CStr(p!ReportID), entries
            Set entries = proposals(CStr(p!ReportID))
            entries.Add Array(p!FieldKey.Value, p!ValueText.Value, p!IsNullValue.Value)
            p.MoveNext
        Loop
        p.Close
    End If
    cacheWorkspace.BeginTrans: trans = True
    Do While Not r.EOF
        cacheRows.AddNew: id = r!ReportID
        rowStates(CStr(id)) = Array(CBool(r!IsArchived), CStr(r!WorkflowState), CLng(Nz(r!CurrentRevisionID, 0)))
        cacheRows!ReportID = id: cacheRows!SortOrdinal = r!SortOrdinal: cacheRows!RevisionID = r!RevisionID
        memberLabel = UXIdentity(UXGridSavedValue(savedValues, r!RevisionID, "OfficialName"), UXGridSavedValue(savedValues, r!RevisionID, "MemberName"))
        cacheRows!Member = memberLabel
        If liveGroup > 0 Then
            cacheRows!ObservationStatus = r!ObservationStatus
            For block = 33 To 39: cacheRows.Fields("Trait" & block) = r.Fields("Trait" & block): Next
            cacheRows!PromotionRecommendation = r!PromotionRecommendation
            cacheRows!AssignmentRecommendationA = r!AssignmentRecommendationA
            cacheRows!AssignmentRecommendationB = r!AssignmentRecommendationB
        Else
            savedValue = UXGridSavedValue(savedValues, r!RevisionID, "ObservationStatus"): cacheRows!ObservationStatus = savedValue
            For block = 33 To 39
                savedValue = UXGridSavedValue(savedValues, r!RevisionID, "Trait" & block)
                cacheRows.Fields("Trait" & block) = savedValue
            Next
            savedValue = UXGridSavedValue(savedValues, r!RevisionID, "PromotionRecommendation"): cacheRows!PromotionRecommendation = savedValue
            savedValue = UXGridSavedValue(savedValues, r!RevisionID, "AssignmentRecommendationA"): cacheRows!AssignmentRecommendationA = savedValue
            savedValue = UXGridSavedValue(savedValues, r!RevisionID, "AssignmentRecommendationB"): cacheRows!AssignmentRecommendationB = savedValue
        End If
        Set previous = priorRows(CStr(id))
        cacheRows!PriorITA = previous("ITA"): cacheRows!PriorPromotion = previous("Promotion"): cacheRows!PriorPeriod = previous("Period"): cacheRows!PriorSource = previous("Source")
        If proposals.Exists(CStr(id)) Then
            Set entries = proposals(CStr(id))
            For Each entry In entries
                'The proposal list supports every canonical field; this cache
                'projects only the visible trait/career columns. Other proposals
                'remain available in the list/detail and normal Apply path.
                Dim projected As DAO.Field
                For Each projected In cacheRows.Fields
                    If projected.Name = CStr(entry(0)) Then
                        If entry(2) Then projected.Value = Null Else projected.Value = entry(1)
                        Exit For
                    End If
                Next
            Next
        End If
        'Set the row ITA inside the cache transaction so the averages pass need not rewrite it.
        cacheRows!ITA = UXGridRowITA(cacheRows)
        cacheRows!RowStatus = UXGridLockReason(id, r!RevisionID)
        If Len(cacheRows!RowStatus) = 0 Then cacheRows!RowStatus = "Editable - saves automatically"
        cacheRows.Update: r.MoveNext
    Loop
    r.Close
    cacheWorkspace.CommitTrans: trans = False
    cacheRows.Close
    Set cacheRows = cacheDB.OpenRecordset("SELECT * FROM PendingTraits ORDER BY SortOrdinal", dbOpenDynaset)
    If Not cacheRows.EOF Then cacheRows.MoveLast: cacheRows.MoveFirst
    Set f.Recordset = cacheRows
    If f.Recordset.EOF Then
        kind = Nz(MajorScalar("SELECT FormKind FROM SummaryGroup WHERE GroupID=" & Nz(Forms!frmUXComparison!uxGroupContext, 0)), "EVAL")
    Else
        kind = Nz(DLookup("FormKind", "Report", "ReportID=" & Nz(f!ReportID, 0)), "EVAL")
    End If
    For block = 33 To 39
        UXGridHeading f.Controls("heading" & block), UXFieldCaption("Trait" & block, kind)
        f.Controls("Trait" & block).ControlTipText = UXFieldCaption("Trait" & block, kind)
    Next
    UXGridHeading f!headingAssignmentRecommendationA, UXFieldCaption("AssignmentRecommendationA", kind)
    UXGridHeading f!headingAssignmentRecommendationB, UXFieldCaption("AssignmentRecommendationB", kind)
    Call UXGridRefreshAverages: loading = False
    Call UXGridCurrent: Call UXGridViewport: UXGridLoad = True
    Exit Function
Failed:
    Set loadFailure = SwiftFailure("modUXTraitGrid", "UXGridLoad", "Build comparison cache", Err.Number, Err.Description, IIf(Len(Err.Source) = 0, "modUXTraitGrid.UXGridLoad", Err.Source), IIf(trans, "Active", "Not started"))
    Resume CleanFailed
CleanFailed:
    If trans Then SwiftRollback cacheWorkspace, loadFailure
    SwiftClose r, loadFailure, "r": SwiftClose p, loadFailure, "p"
    On Error Resume Next
    Set f.Recordset = Nothing
    If Err.Number <> 0 Then SwiftCleanupError loadFailure, "detach grid", Err.Number
    On Error GoTo 0
    If Not UXGridClose() Then SwiftCleanupError loadFailure, "remove comparison cache", vbObjectError + 4816
    Call SwiftFailureText(loadFailure) 'the load failure, not the cache cleanup, stays the last failure
    loading = False
    Err.Raise CLng(loadFailure("Number")), CStr(loadFailure("Source")), CStr(loadFailure("Description"))
End Function
Private Function UXGridSavedValue(ByVal values As Object, ByVal revision As Long, ByVal key As String) As Variant
    Dim token As String
    token = revision & ":" & key
    If Not values.exists(token) Then Err.Raise vbObjectError + 4007, "modUXTraitGrid.UXGridSavedValue", "Snapshot field missing."
    UXGridSavedValue = values(token)
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
    If state(1) <> "Draft" And state(1) <> "Approved" And state(1) <> "Signed" Then UXGridLockReason = "Review lock - reopen report"
End Function
' Apply cached row locks and contextual status on row focus; avoid whole-group queries.
Public Function UXGridCurrent(Optional ByVal authoritative As Boolean = False) As Boolean
    Dim f As Access.Form, reason As String, block As Long
    If loading Or saving Or cacheRows Is Nothing Then Exit Function
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If UXHasControl(Forms!frmUXComparison, "uxOpenGroupReport") Then Forms!frmUXComparison!uxOpenGroupReport.Enabled = Not f.Recordset.EOF
    If f.Recordset.EOF Then Exit Function
    f!PriorITA.ControlTipText = Nz(f!PriorPeriod, "") & " - " & Nz(f!PriorSource, "")
    f!PriorPromotion.ControlTipText = f!PriorITA.ControlTipText
    reason = UXGridLockReason(f!ReportID, f!RevisionID, authoritative)
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
    Forms!frmUXComparison!lblStatus.Caption = "Not saved yet - " & UXGridPendingCount() & " pending cells; saving automatically..."
    UXGridPendingActions
    UXApplyVisual Forms!frmUXComparison
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
        If CLng(Cell(3)) <> MajorAuthor() Then Err.Raise vbObjectError + 4815, "modUXTraitGrid.UXGridDrain", "Pending cells belong to another contributor. Switch back before retrying or discard these cells."
        Set r = f.RecordsetClone: r.FindFirst "ReportID=" & cell(0)
        If r.NoMatch Then Err.Raise vbObjectError + 4813, "modUXTraitGrid.UXGridDrain", "Pending report is no longer in the comparison."
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
        Forms!frmUXComparison!lblStatus.Caption = "Saved automatically - no pending cells. Local report revisions only."
        UXApplyVisual Forms!frmUXComparison
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
    UXGridPendingActions
    UXApplyVisual Forms!frmUXComparison
    Exit Sub
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
    If Not IsTraitKey(key) And key <> "PromotionRecommendation" And key <> "AssignmentRecommendationA" And key <> "AssignmentRecommendationB" Then Err.Raise vbObjectError + 4812, "modUXTraitGrid.UXGridEdit", "Only traits, promotion and career recommendations can be edited here."
    If key = "AssignmentRecommendationA" Or key = "AssignmentRecommendationB" Then
        If Len(Nz(value, "")) > 20 Then Err.Raise vbObjectError + 4811, "modUXTraitGrid.UXGridEdit", "Career recommendations allow up to 20 characters."
    End If
    stage = "Check edit guards"
    reason = UXGridLockReason(f!ReportID, f!RevisionID, True)
    If key = "PromotionRecommendation" Then
        Dim validated As Variant
        validated = ValidateAssessment(Array(f!ObservationStatus.Value, value))
    End If
    If Len(reason) > 0 Then Err.Raise vbObjectError + 4810, "modUXTraitGrid.UXGridEdit", reason
    If Left$(key, 5) = "Trait" And Not IsNull(value) Then
        If Len(value) > 0 And InStr(1, ";NOB;1.0;2.0;3.0;4.0;5.0;", ";" & value & ";", vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 4811, "modUXTraitGrid.UXGridEdit", "Choose NOB or a whole mark from 1.0 to 5.0."
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
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "") = "Signed" Then
        If reopenDecision = 0 Then reopenDecision = MsgBox("This report was marked signed. Reopen it for correction and save this change? The signed version stays in history.", vbYesNo + vbQuestion + vbDefaultButton2, "Reopen signed report")
        If reopenDecision <> vbYes Then
            f.Undo: saving = False: UXGridEdit = False: Exit Function
        End If
    ElseIf Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "") = "Approved" Then
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
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "") = "Signed" Then ReopenSignedReport reportID, "Group members change to " & key
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
    Forms!frmUXComparison!lblStatus.Caption = "Saved automatically at " & Format$(Now(), "h:nn:ss AM/PM") & " - local report revision"
    saving = False: UXGridEdit = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXTraitGrid", "UXGridEdit", stage, Err.number, Err.description, Err.source, "Unknown")
    failure = stage & ": " & CLng(swiftPrimary("Number")) & " " & CStr(swiftPrimary("Description"))
    If started Then SwiftRollback w, swiftPrimary
    saving = False
    If committed Then
        swiftPrimary("Transaction") = "Committed"
        Forms!frmUXComparison!lblStatus.Caption = "Report saved, but the comparison refresh failed: " & SwiftFailureText(swiftPrimary) & ". Close and reopen the comparison."
        UXGridEdit = True: Exit Function
    End If
    On Error Resume Next
    f.Undo
    If baselineRead Then
        f.Controls(key) = before: f!RevisionID = expectedRevision
        If f.Dirty Then f.Dirty = False
        UXGridRefreshAverages
    End If
    If Err.Number <> 0 Then SwiftCleanupError swiftPrimary, "restore displayed cell", Err.Number
    On Error GoTo 0
    Forms!frmUXComparison!lblStatus.Caption = SwiftFailureText(swiftPrimary) & " Cell change was not saved."
End Function
' Compute pending cache ITAs and the group trait average, excluding wholly NOB reports.
Private Function UXGridRowTotals(ByVal r As DAO.Recordset, ByRef rowSum As Double) As Long
    Dim block As Long, value As String
    rowSum = 0
    If Nz(r!ObservationStatus, "") = "NOB" Then Exit Function
    For block = 33 To 39
        value = Nz(r.Fields("Trait" & block), "")
        If value <> "NOB" And Len(value) > 0 And IsNumeric(value) Then UXGridRowTotals = UXGridRowTotals + 1: rowSum = rowSum + CDbl(value)
    Next
End Function
Private Function UXGridRowITA(ByVal r As DAO.Recordset) As String
    Dim rowCount As Long, rowSum As Double
    rowCount = UXGridRowTotals(r, rowSum)
    If rowCount > 0 Then UXGridRowITA = Format$(rowSum / rowCount, "0.00") Else UXGridRowITA = "-"
End Function
' Write an ITA cache row only when its formatted value changed.
Public Function UXGridRefreshAverages() As Boolean
    Dim r As DAO.Recordset, count As Long, sum As Double, rowCount As Long, rowSum As Double, nob As Long, average As String
    Set r = cacheDB.OpenRecordset("PendingTraits", dbOpenDynaset)
    Do While Not r.EOF
        rowCount = UXGridRowTotals(r, rowSum)
        If Nz(r!ObservationStatus, "") = "NOB" Then nob = nob + 1
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
    f!lstProposals.RowSource = "SELECT p.ProposalID,UXIdentity(r.OfficialName,r.MemberName) AS Member,UXFieldCaption(p.FieldKey,r.FormKind) AS Block,MajorRevisionValue(p.BaseRevisionID,p.FieldKey) AS [Before],p.ValueText AS [After],p.ProposalState AS State FROM qryMeetingProposalRevisionIDs AS p INNER JOIN Report AS r ON p.ReportID=r.ReportID WHERE p.SetID=" & f!txtMeeting & " AND p.ProposalState IN (" & IIf(results, "'Proposed','Applied'", "'Proposed'") & ") ORDER BY UXIdentity(r.OfficialName,r.MemberName),p.FieldKey"
    f!lstProposals.Requery: UXRefreshComparisonPreview = True
End Function
' Show the selected proposal's complete before/after values outside the condensed list.
Public Function UXGridDetail() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, db As DAO.Database
    Set f = Forms!frmUXComparison
    If Nz(f!lstProposals, 0) = 0 Then Exit Function
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM qryMeetingProposalRevisionIDs WHERE ProposalID=" & f!lstProposals, dbOpenSnapshot)
    f!uxProposalDetail = UXFieldCaption(r!FieldKey, Nz(DLookup("FormKind", "Report", "ReportID=" & r!ReportID), "EVAL")) & vbCrLf & "BEFORE: " & Nz(MajorRevisionValue(r!BaseRevisionID, r!FieldKey), "(blank)") & vbCrLf & "AFTER: " & Nz(r!ValueText, "(clear value)")
    r.Close: UXGridDetail = True
End Function

' Fit the native grid columns to the viewport and retain the live or historical group context.
Public Function UXGridViewport() As Boolean
    Dim f As Access.Form, c As Access.Control, v As Variant, usable As Long, ratio As Double
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    UXResizeWorkspace "frmUXComparison"
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    'Shrink columns to fit a narrow window; wider windows leave the designed widths.
    usable = Forms!frmUXComparison!uxTraitGrid.Width - 350: ratio = usable / 17050: If ratio > 1 Then ratio = 1
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

Public Function UXSortGrid(ByVal key As String) As Boolean
 Dim f As Access.Form, field As DAO.Field, allowed As Boolean, id As Long, rows As DAO.Recordset
 On Error GoTo Failed
 If Not UXGridDrain() Then Exit Function
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 If cacheDB Is Nothing Then Exit Function
 If f.Recordset.EOF Then UXSortGrid = True: Exit Function
 For Each field In cacheDB.TableDefs("PendingTraits").Fields
  If field.Name = key Then allowed = True
 Next
 If Not allowed Then Exit Function
 id = Nz(f!ReportID, 0)
 If f.Dirty Then f.Dirty = False
 If gridSortKey = key Then gridSortDescending = Not gridSortDescending Else gridSortDescending = False
 gridSortKey = key
 Set rows = cacheDB.OpenRecordset("SELECT * FROM PendingTraits ORDER BY [" & key & "] " & IIf(gridSortDescending, "DESC", "ASC") & ", SortOrdinal", dbOpenDynaset)
 Set f.Recordset = rows
 If id > 0 Then rows.FindFirst "ReportID=" & id
 UXSortGrid = True: Exit Function
Failed:
 Forms!frmUXComparison!lblStatus.Caption = "Unable to sort: " & Err.Description & ". Pending edits were retained."
End Function

'Compare displayed numeric ITAs only; missing, NOB and invalid values stay neutral.
Public Function UXITAChange(ByVal current As Variant, ByVal prior As Variant) As Long
 If IsNull(current) Or IsNull(prior) Then Exit Function
 If Not IsNumeric(current) Or Not IsNumeric(prior) Then Exit Function
 If CDbl(current) < 1 Or CDbl(current) > 5 Or CDbl(prior) < 1 Or CDbl(prior) > 5 Then Exit Function
 If CDbl(current) > CDbl(prior) Then UXITAChange = 1
 If CDbl(current) < CDbl(prior) Then UXITAChange = -1
End Function
'Grid columns are too narrow for the official captions; show a short heading and
'keep the official caption as the tooltip. The editor pages keep official captions.
Private Sub UXGridHeading(ByVal heading As Access.Control, ByVal caption As String)
    Dim words As Variant, i As Long, shown As String, word As String
    heading.ControlTipText = caption
    caption = Replace(Replace(caption, "  ", " "), "/", "/ ")
    caption = Replace(caption, "Career Recommendations - first", "Career rec. 1")
    caption = Replace(caption, "Career Recommendations - second", "Career rec. 2")
    words = Split(caption, " ")
    For i = 0 To UBound(words)
        word = words(i)
        If i = 0 Then
            word = Replace(word, ".", "")
        Else
            Select Case UCase$(Replace(word, "/", ""))
                Case "PROFESSIONAL": word = "Prof."
                Case "KNOWLEDGE": word = "knowl."
                Case "EXPERTISE": word = "expert."
                Case "TECHNICAL": word = "Tech."
                Case "COMMAND": word = "Cmd."
                Case "ORGANIZATIONAL": word = "org."
                Case "INSTITUTIONAL": word = "Instit."
                Case "MILITARY": word = "Mil."
                Case "CHARACTER": word = "char."
                Case "PERSONAL": word = "Pers."
                Case "ACCOMPLISHMENT": word = "accomp."
                Case "INITIATIVE": word = "init."
                Case "PERFORMANCE": word = "perf."
                Case "EFFECTIVENESS": word = "effect."
                Case "AND": word = "&&"
                Case "PROFESSIONALISM": word = "Profes-" & vbCrLf & "sional-" & vbCrLf & "ism"
                Case "ACCOUNTABILITY": word = "Account-" & vbCrLf & "ability"
                Case "TEAMWORK": word = "Team-" & vbCrLf & "work"
                Case "LEADERSHIP": word = "Leader-" & vbCrLf & "ship"
                Case "DECKPLATE": word = "Deck-" & vbCrLf & "plate"
                Case Else: word = IIf(i = 1, UCase$(Left$(word, 1)) & LCase$(Mid$(word, 2)), LCase$(word))
            End Select
            If Right$(words(i), 1) = "/" And Right$(word, 1) <> "/" Then word = word & "/"
        End If
        shown = shown & IIf(i = 0, "", IIf(Right$(shown, 1) = vbLf, "", " ")) & word
    Next
    heading.Caption = shown
End Sub
