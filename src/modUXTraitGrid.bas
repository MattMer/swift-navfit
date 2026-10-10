Attribute VB_Name = "modUXTraitGrid"
Option Explicit
Private cacheWorkspace As DAO.Workspace
Private cacheDB As DAO.Database
Private cacheRows As DAO.Recordset
Private cachePath As String
Private loading As Boolean
Private saving As Boolean
Private queuedCells As Object
Public Function UXGridClose() As Boolean
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
Public Function UXGridLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, p As DAO.Recordset, db As DAO.Database, sql As String, block As Long, id As Long, kind As String, previous As Object
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
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM ComparisonMember WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " ORDER BY SortOrdinal", dbOpenSnapshot)
    Do While Not r.EOF
        cacheRows.AddNew: id = r!ReportID
        cacheRows!ReportID = id: cacheRows!SortOrdinal = r!SortOrdinal: cacheRows!RevisionID = r!RevisionID
        cacheRows!Member = UXRevisionMemberLabel(r!RevisionID): cacheRows!ObservationStatus = MajorRevisionValue(r!RevisionID, "ObservationStatus")
        For block = 33 To 39: cacheRows.Fields("Trait" & block) = MajorRevisionValue(r!RevisionID, "Trait" & block): Next
        cacheRows!PromotionRecommendation = MajorRevisionValue(r!RevisionID, "PromotionRecommendation")
        cacheRows!AssignmentRecommendationA = MajorRevisionValue(r!RevisionID, "AssignmentRecommendationA")
        cacheRows!AssignmentRecommendationB = MajorRevisionValue(r!RevisionID, "AssignmentRecommendationB")
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
Public Function UXGridLockReason(ByVal id As Long, ByVal revision As Long) As String
    Dim snapshot As Long, groupID As Long
    If CurrentRevision(id) <> revision Then UXGridLockReason = "Saved report changed - reopen comparison": Exit Function
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then UXGridLockReason = "Archived - recover report to edit": Exit Function
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" And Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Approved" Then UXGridLockReason = "Review lock - reopen report": Exit Function
    snapshot = Nz(DLookup("GroupSnapshotID", "ComparisonSet", "SetID=" & Forms!frmUXComparison!txtMeeting), 0)
    groupID = Nz(DLookup("GroupID", "GroupSnapshot", "SnapshotID=" & snapshot), 0)
    If Nz(DLookup("Stamp", "GroupSnapshot", "SnapshotID=" & snapshot), "") <> GroupStamp(groupID) Then UXGridLockReason = "Group changed - reopen comparison"
End Function
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
    Forms!frmUXComparison!lblStatus.Caption = IIf(Len(reason) > 0, reason, "Edit cells directly. Valid changes save automatically as report revisions.")
    UXGridCurrent = True
End Function
Public Function UXGridQueue(ByVal key As String) As Boolean
    Dim f As Access.Form, identity As String
    On Error GoTo Failed
    If loading Or saving Then Exit Function
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If queuedCells Is Nothing Then Set queuedCells = CreateObject("Scripting.Dictionary")
    identity = CStr(f!ReportID) & ":" & key
    queuedCells(identity) = Array(CLng(f!ReportID), key, f.Controls(key).Value)
    Forms!frmUXComparison.TimerInterval = 150
    Forms!frmUXComparison!lblStatus.Caption = "Saving pending cell changes..."
    UXGridQueue = True: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = "Queue cell save: " & Err.Number & " " & Err.Description
End Function
Public Function UXGridDrain() As Boolean
    Dim pending As Object, item As Variant, cell As Variant, f As Access.Form, r As DAO.Recordset, ok As Boolean, failure As String
    On Error GoTo Failed
    If queuedCells Is Nothing Or saving Then UXGridDrain = True: Exit Function
    Forms!frmUXComparison.TimerInterval = 0
    Set pending = queuedCells: Set queuedCells = Nothing
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    'Finish Access's bound-cache update before entering any report transaction.
    If f.Dirty Then f.Dirty = False
    ok = True
    For Each item In pending.Keys
        cell = pending(item)
        Set r = f.RecordsetClone: r.FindFirst "ReportID=" & cell(0)
        If r.NoMatch Then Err.Raise vbObjectError + 4813, , "Pending report is no longer in the comparison."
        f.Bookmark = r.Bookmark: r.Close
        f.Controls(CStr(cell(1))) = cell(2)
        If Not UXGridEdit(CStr(cell(1))) Then ok = False: Exit For
    Next
    If Not ok Then
        failure = Forms!frmUXComparison!lblStatus.Caption
        UXGridLoad
        Forms!frmUXComparison!lblStatus.Caption = failure
    End If
    UXGridDrain = ok: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = "Pending cell save: " & Err.Number & " " & Err.Description
End Function
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
    reason = UXGridLockReason(f!ReportID, f!RevisionID)
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
    snapshot = Nz(DLookup("GroupSnapshotID", "ComparisonSet", "SetID=" & Forms!frmUXComparison!txtMeeting), 0)
    groupID = Nz(DLookup("GroupID", "GroupSnapshot", "SnapshotID=" & snapshot), 0)
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
    stage = "Refresh group baseline"
    snapshot = FreezeSummaryGroup(groupID, "Comparison refreshed after automatic cell save")
    stage = "Refresh comparison baseline"
    newSet = CreateComparisonSet(snapshot, "Group members - automatic save")
    stage = "Flush edited comparison row"
    If f.Dirty Then f.Dirty = False
    w.CommitTrans: started = False: committed = True
    stage = "Refresh saved row"
    Forms!frmUXComparison!txtMeeting = newSet
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
        r.Edit
        If rowCount > 0 Then r!ITA = Format$(rowSum / rowCount, "0.00") Else r!ITA = "-"
        r.Update: sum = sum + rowSum: count = count + rowCount: r.MoveNext
    Loop
    r.Close
    Forms!frmUXComparison!uxGridAverage.Caption = "Group trait average: " & IIf(count > 0, Format$(sum / IIf(count = 0, 1, count), "0.00"), "Unavailable") & " | " & count & " numeric traits | " & nob & " whole-report NOB excluded"
    Forms!frmUXComparison!uxTraitGrid.Form.Refresh
    UXGridRefreshAverages = True
End Function
Public Function UXRefreshComparisonPreview(Optional ByVal results As Boolean = False) As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXComparison
    f!lstProposals.RowSource = "SELECT p.ProposalID,UXIdentity(r.OfficialName,r.MemberName) AS Member,UXFieldCaption(p.FieldKey,r.FormKind) AS Block,MajorRevisionValue(p.BaseRevision,p.FieldKey) AS [Before],p.ValueText AS [After],p.ProposalState AS State FROM MeetingProposal AS p INNER JOIN Report AS r ON p.ReportID=r.ReportID WHERE p.SetID=" & f!txtMeeting & " AND p.ProposalState IN (" & IIf(results, "'Proposed','Applied'", "'Proposed'") & ") ORDER BY UXIdentity(r.OfficialName,r.MemberName),p.FieldKey"
    f!lstProposals.Requery: UXRefreshComparisonPreview = True
End Function
Public Function UXGridDetail() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, db As DAO.Database
    Set f = Forms!frmUXComparison
    If Nz(f!lstProposals, 0) = 0 Then Exit Function
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE ProposalID=" & f!lstProposals, dbOpenSnapshot)
    f!uxProposalDetail = UXFieldCaption(r!FieldKey, Nz(DLookup("FormKind", "Report", "ReportID=" & r!ReportID), "EVAL")) & vbCrLf & "BEFORE: " & Nz(MajorRevisionValue(r!BaseRevision, r!FieldKey), "(blank)") & vbCrLf & "AFTER: " & Nz(r!ValueText, "(clear value)")
    r.Close: UXGridDetail = True
End Function

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
    Forms!frmUXComparison!uxGroupContext = Nz(MajorScalar("SELECT g.GroupID FROM GroupSnapshot AS g INNER JOIN ComparisonSet AS c ON g.SnapshotID=c.GroupSnapshotID WHERE c.SetID=" & Nz(Forms!frmUXComparison!txtMeeting, 0)), 0)
    UXGridViewport = True
End Function

Public Function UXGridUnload(ByVal f As Access.Form) As Boolean
    UXGridDrain
    loading = True
    Set f!uxTraitGrid.Form.Recordset = Nothing
    UXGridClose
    UXGridUnload = True
End Function
Public Function UXGridCacheFile() As String
    UXGridCacheFile = cachePath
End Function

Public Function UXCloseComparison() As Boolean
    DoCmd.Close acForm, "frmUXComparison", acSaveNo
    If CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then
        Forms!frmMajorWorkbench.Visible = True
        Forms!frmMajorWorkbench.SetFocus
    Else
        OpenScreen "frmMajorWorkbench"
    End If
    UXCloseComparison = True
End Function
