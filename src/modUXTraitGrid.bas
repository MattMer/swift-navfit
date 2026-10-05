Attribute VB_Name = "modUXTraitGrid"
Option Explicit
Private cacheDB As DAO.Database
Private cacheRows As DAO.Recordset
Private cachePath As String
Private loading As Boolean
Public Function UXGridClose() As Boolean
    loading = True
    On Error Resume Next
    If Not cacheRows Is Nothing Then cacheRows.Close
    Set cacheRows = Nothing
    If Not cacheDB Is Nothing Then cacheDB.Close
    Set cacheDB = Nothing
    If Len(cachePath) > 0 Then Kill cachePath
    If Len(cachePath) > 0 Then
        If Len(Dir$(cachePath)) = 0 Then cachePath = ""
    End If
    UXGridClose = True
End Function
Public Function UXGridLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, p As DAO.Recordset, db As DAO.Database, sql As String, block As Long, id As Long, kind As String
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    loading = True
    Set f.Recordset = Nothing
    UXGridClose
    cachePath = Environ$("TEMP") & "\Swift-NAVFIT-traits-" & Replace(NewIdentity(), "-", "") & ".accdb"
    Set cacheDB = DBEngine.CreateDatabase(cachePath, dbLangGeneral, dbVersion120)
    sql = "CREATE TABLE PendingTraits (ReportID LONG PRIMARY KEY, SortOrdinal LONG, RevisionID LONG, Member TEXT(120), RowStatus TEXT(120), ObservationStatus TEXT(8), ITA TEXT(12)"
    For block = 33 To 39: sql = sql & ", Trait" & block & " TEXT(3)": Next
    sql = sql & ", PromotionRecommendation TEXT(20))": cacheDB.Execute sql, dbFailOnError
    Set cacheRows = cacheDB.OpenRecordset("PendingTraits", dbOpenDynaset)
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM ComparisonMember WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " ORDER BY SortOrdinal", dbOpenSnapshot)
    Do While Not r.EOF
        cacheRows.AddNew: id = r!ReportID
        cacheRows!ReportID = id: cacheRows!SortOrdinal = r!SortOrdinal: cacheRows!RevisionID = r!RevisionID
        cacheRows!Member = UXRevisionMemberLabel(r!RevisionID): cacheRows!ObservationStatus = MajorRevisionValue(r!RevisionID, "ObservationStatus")
        For block = 33 To 39: cacheRows.Fields("Trait" & block) = MajorRevisionValue(r!RevisionID, "Trait" & block): Next
        cacheRows!PromotionRecommendation = MajorRevisionValue(r!RevisionID, "PromotionRecommendation")
        Set p = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " AND ReportID=" & id & " AND ProposalState='Proposed' ORDER BY ProposalID", dbOpenSnapshot)
        Do While Not p.EOF
            If p!IsNullValue Then cacheRows.Fields(p!FieldKey) = Null Else cacheRows.Fields(p!FieldKey) = p!ValueText
            p.MoveNext
        Loop
        p.Close
        cacheRows!RowStatus = UXGridLockReason(id, r!RevisionID)
        If Len(cacheRows!RowStatus) = 0 Then cacheRows!RowStatus = "Editable - proposals only"
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
    Call UXGridRefreshAverages: loading = False
    Call UXGridCurrent: Call UXGridViewport: UXGridLoad = True
End Function
Public Function UXGridLockReason(ByVal id As Long, ByVal revision As Long) As String
    Dim snapshot As Long, groupID As Long
    If CurrentRevision(id) <> revision Then UXGridLockReason = "Saved report changed - reopen comparison": Exit Function
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then UXGridLockReason = "Archived - recover report to edit": Exit Function
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then UXGridLockReason = "Local approval / review lock - reopen report": Exit Function
    snapshot = Nz(DLookup("GroupSnapshotID", "ComparisonSet", "SetID=" & Forms!frmUXComparison!txtMeeting), 0)
    groupID = Nz(DLookup("GroupID", "GroupSnapshot", "SnapshotID=" & snapshot), 0)
    If Nz(DLookup("Stamp", "GroupSnapshot", "SnapshotID=" & snapshot), "") <> GroupStamp(groupID) Then UXGridLockReason = "Group changed - reopen comparison"
End Function
Public Function UXGridCurrent() As Boolean
    Dim f As Access.Form, reason As String, block As Long
    If loading Or cacheRows Is Nothing Then Exit Function
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then Exit Function
    Set f = Forms!frmUXComparison!uxTraitGrid.Form
    If f.Recordset.EOF Then Exit Function
    reason = UXGridLockReason(f!ReportID, f!RevisionID)
    For block = 33 To 39: f.Controls("Trait" & block).Locked = (Len(reason) > 0): Next
    f!PromotionRecommendation.Locked = (Len(reason) > 0)
    Forms!frmUXComparison!lblStatus.Caption = IIf(Len(reason) > 0, reason, "Edit cells directly. Changes are staged; Preview and Apply save reviewed revisions.")
    UXGridCurrent = True
End Function
Public Function UXGridEdit(ByVal key As String) As Boolean
    Dim f As Access.Form, value As Variant, reason As String, w As DAO.Workspace, started As Boolean, db As DAO.Database, before As Variant, proposal As Long, failure As String
    On Error GoTo Failed
    Set f = Forms!frmUXComparison!uxTraitGrid.Form: value = f.Controls(key).Value
    If Not IsTraitKey(key) And key <> "PromotionRecommendation" Then Err.Raise vbObjectError + 4812, , "Only trait marks and promotion recommendation can be edited here."
    reason = UXGridLockReason(f!ReportID, f!RevisionID)
    If key = "PromotionRecommendation" Then
        Dim validated As Variant
        validated = ValidateAssessment(Array(f!ObservationStatus.Value, value))
    End If
    If Len(reason) > 0 Then Err.Raise vbObjectError + 4810, , reason
    If Left$(key, 5) = "Trait" And Not IsNull(value) Then
        If Len(value) > 0 And InStr(1, ";NOB;1.0;2.0;3.0;4.0;5.0;", ";" & value & ";", vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 4811, , "Choose NOB or a whole mark from 1.0 to 5.0."
    End If
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    w.BeginTrans: started = True
    'Supersede only this UI session's current field proposal. Prior values remain auditable.
    db.Execute "UPDATE MeetingProposal SET ProposalState='Superseded' WHERE SetID=" & Forms!frmUXComparison!txtMeeting & " AND ReportID=" & f!ReportID & " AND FieldKey=" & QuoteSQL(key) & " AND ProposalState='Proposed'", dbFailOnError
    before = MajorRevisionValue(f!RevisionID, key)
    If Not SameFieldValue(before, value) Then proposal = ProposeMeetingField(Forms!frmUXComparison!txtMeeting, f!ReportID, key, value)
    If f.Dirty Then f.Dirty = False
    w.CommitTrans: started = False
    UXGridRefreshAverages
    UXRefreshComparisonPreview
    Forms!frmUXComparison!lblStatus.Caption = "Cell change staged. Saved reports are unchanged; preview before applying."
    UXGridEdit = True: Exit Function
Failed:
    failure = Err.Description
    If started Then w.Rollback
    On Error Resume Next
    f.Undo
    Forms!frmUXComparison!lblStatus.Caption = failure & " Cell change was not staged."
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
    Forms!frmUXComparison!uxGridAverage.Caption = "Pending group trait average: " & IIf(count > 0, Format$(sum / IIf(count = 0, 1, count), "0.00"), "Unavailable") & " | " & count & " numeric traits | " & nob & " whole-report NOB excluded"
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
    usable = Forms!frmUXComparison!uxTraitGrid.Width - 350: ratio = usable / 10450
    For Each c In f.Controls
        If InStr(c.Tag, ":") > 0 Then
            v = Split(c.Tag, ":"): c.Left = CLng(v(0) * ratio): c.Width = CLng(v(1) * ratio)
        End If
    Next
    f.Width = usable: f.Section(acHeader).Height = IIf(usable < 9000, 1500, 950)
    UXGridViewport = True
End Function

Public Function UXGridUnload(ByVal f As Access.Form) As Boolean
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
        Forms!frmMajorWorkbench.SetFocus
    Else
        OpenScreen "frmMajorWorkbench"
    End If
    UXCloseComparison = True
End Function
