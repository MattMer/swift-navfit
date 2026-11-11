Attribute VB_Name = "modUXDirectory"
Option Explicit
Private memberID As Long, memberVersion As Long
Private historyMember As Long, priorID As Long, priorVersion As Long
Private counselingID As Long, counselingVersion As Long, counselingMember As Long
Public Function UXDirectoryLoad() As Boolean
    Dim f As Access.Form, term As String
    RepairPDFDirectoryLinks
    Set f = Forms!frmUXDirectory: term = Replace(Nz(f!txtSearch, ""), "'", "''")
    f!lstMembers.RowSourceType = "Table/Query"
    'Official report name (PARKER, REESE) when entered; search finds either name.
    f!lstMembers.RowSource = "SELECT MemberID,IIf(Nz(ReportName,'')='',FullName,ReportName) AS [Member],RankRate AS [Rank/rate],IIf(MemberSSN Is Null,'Unentered','***-**-' & Right(MemberSSN,4)) AS [SSN],SourceKind AS [Source] FROM CommandMember WHERE IsArchived=False AND (InStr(1,FullName,'" & term & "',1)>0 OR InStr(1,Nz(ReportName,''),'" & term & "',1)>0) ORDER BY IIf(Nz(ReportName,'')='',FullName,ReportName)"
    f!lstMembers.Requery: UXDirectoryLoad = True
End Function
Public Function UXDirectoryEdit(Optional ByVal createNew As Boolean = False) As Boolean
    Dim r As DAO.Recordset, f As Access.Form, db As DAO.Database
    If Not createNew And IsNull(Forms!frmUXDirectory!lstMembers) Then Exit Function
    memberID = 0: memberVersion = 0
    If Not createNew Then memberID = CLng(Forms!frmUXDirectory!lstMembers)
    DoCmd.OpenForm "frmUXMemberEditor": Set f = Forms!frmUXMemberEditor
    If memberID > 0 Then
        Set db = CurrentDb
        Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
        memberVersion = r!VersionCounter: f!txtName = r!FullName: f!txtReportName = r!ReportName: f!txtRank = r!RankRate: f!txtSSN = r!MemberSSN: f!txtDateReported = IIf(IsNull(r!DateReported), Null, Format$(r!DateReported, "yyyy-mm-dd")): UXRUADLoad r!RUADData: r.Close
    Else
        UXRUADLoad Null
    End If
    UXDirectoryEdit = True
End Function
Public Function UXDirectorySave() As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXMemberEditor
    memberID = SaveDirectoryMember(memberID, memberVersion, Nz(f!txtName, ""), Nz(f!txtReportName, ""), Nz(f!txtRank, ""), Nz(f!txtSSN, ""), "Manual", UXRUADValues(), , Nz(f!txtDateReported, ""), True)
    UXDirectorySave = True
    If CurrentProject.AllForms("frmUXDirectory").IsLoaded Then UXDirectoryLoad
    'Stay open to say which open reports took the change and which did not.
    If Len(LastMemberReportUpdate()) > 0 Then
        memberVersion = Nz(MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & memberID), 0)
        f!lblStatus.Caption = Left$("Saved. " & LastMemberReportUpdate(), 1000)
        UXEditorContextChanged
        Exit Function
    End If
    DoCmd.Close acForm, f.Name: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXDirectoryClose(ByVal screen As String) As Boolean
    If screen = "frmUXDirectoryBatch" Then ClearDirectoryReportPreview
    DoCmd.Close acForm, screen: UXDirectoryClose = True
End Function
Public Function UXDirectoryHistory() As Boolean
    If IsNull(Forms!frmUXDirectory!lstMembers) Then Exit Function
    historyMember = CLng(Forms!frmUXDirectory!lstMembers)
    DoCmd.OpenForm "frmUXMemberHistory": UXHistoryLoad: UXDirectoryHistory = True
End Function
Public Function UXHistoryLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, id As Long, priorITAs As Object, swiftITAs As Object, historyText As String
    Set f = Forms!frmUXMemberHistory
    f!lblMember.Caption = Nz(MajorScalar("SELECT FullName FROM CommandMember WHERE MemberID=" & historyMember), "")
    f!lstHistory.RowSourceType = "Value List": f!lstHistory.RowSource = ""
    Set priorITAs = PriorEvaluationITARows(historyMember): Set swiftITAs = UXHistoryITARows(historyMember)
    Set r = CurrentDb.OpenRecordset("SELECT PriorID,FormKind,PeriodStart,PeriodEnd,PromotionRecommendation FROM PriorEvaluation WHERE MemberID=" & historyMember & " AND IsArchived=False ORDER BY PeriodEnd DESC", dbOpenSnapshot)
    Do While Not r.EOF
        If historyText <> "" Then historyText = historyText & ";"
        id = r!PriorID: historyText = historyText & CStr(id) & ";Manual prior;" & r!FormKind & ";" & Format$(r!PeriodStart, "yyyy-mm-dd") & " to " & Format$(r!PeriodEnd, "yyyy-mm-dd") & ";" & priorITAs(CStr(id)) & ";" & Nz(r!PromotionRecommendation, ""): r.MoveNext
    Loop: r.Close
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID,r.FormKind,r.PeriodStart,r.PeriodEnd,r.PromotionRecommendation,m.MatchState FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & historyMember & " ORDER BY r.PeriodEnd DESC", dbOpenSnapshot)
    Do While Not r.EOF
        If historyText <> "" Then historyText = historyText & ";"
        id = r!ReportID: historyText = historyText & CStr(-id) & ";Swift - " & r!MatchState & ";" & r!FormKind & ";" & Format$(Nz(r!PeriodStart, ""), "yyyy-mm-dd") & " to " & Format$(Nz(r!PeriodEnd, ""), "yyyy-mm-dd") & ";" & swiftITAs(CStr(id)) & ";" & Nz(r!PromotionRecommendation, ""): r.MoveNext
    Loop: r.Close
    f!lstHistory.RowSource = historyText
    UXHistoryCounselingLoad
    UXHistoryLoad = True
End Function
Public Function UXPriorOpen(Optional ByVal createNew As Boolean = True) As Boolean
    Dim f As Access.Form, r As DAO.Recordset
    priorID = 0: priorVersion = 0
    If Not createNew Then
        If IsNull(Forms!frmUXMemberHistory!lstHistory) Then Exit Function
        priorID = CLng(Forms!frmUXMemberHistory!lstHistory)
        If priorID < 0 Then Forms!frmUXMemberHistory!lblStatus.Caption = "Saved Swift reports are maintained in Reports. Only manually recorded prior evaluations are edited here.": Exit Function
    End If
    DoCmd.OpenForm "frmUXPriorEditor": Set f = Forms!frmUXPriorEditor
    f!cboKind = "EVAL": f!cboObservation = "Observed": f!txtRank = MajorScalar("SELECT RankRate FROM CommandMember WHERE MemberID=" & historyMember)
    If priorID > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & priorID, dbOpenSnapshot)
        priorVersion = r!VersionCounter: f!cboKind = r!FormKind: f!txtRank = r!RankRate: f!txtFrom = r!PeriodStart: f!txtTo = r!PeriodEnd: f!cboPromotion = r!PromotionRecommendation: f!cboObservation = r!ObservationStatus: f!txtSource = r!SourceNote: r.Close
        Set r = CurrentDb.OpenRecordset("SELECT BlockNumber,MarkValue FROM PriorEvaluationTrait WHERE PriorID=" & priorID, dbOpenSnapshot)
        Do While Not r.EOF: f.Controls("cboTrait" & r!BlockNumber) = r!MarkValue: r.MoveNext: Loop: r.Close
    End If
    Call UXPriorLabels: UXPriorOpen = True
End Function
Public Function UXPriorLabels() As Boolean
    Dim block As Long, f As Access.Form
    Set f = Forms!frmUXPriorEditor
    For block = 33 To 39: f.Controls("lblTrait" & block).Caption = UXFieldCaption("Trait" & block, Nz(f!cboKind, "EVAL")): Next
    UXPriorLabels = True
End Function
Public Function UXPriorSave() As Boolean
    Dim f As Access.Form, traits As Object, block As Long, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXPriorEditor
    If Not IsDate(f!txtFrom) Or Not IsDate(f!txtTo) Then Err.Raise vbObjectError + 6975, "modUXDirectory.UXPriorSave", "Enter both report dates."
    Set traits = CreateObject("Scripting.Dictionary")
    For block = 33 To 39: traits.Add "Trait" & block, Nz(f.Controls("cboTrait" & block), ""): Next
    id = SavePriorEvaluation(priorID, priorVersion, historyMember, Nz(f!cboKind, ""), Nz(f!txtRank, ""), CDate(f!txtFrom), CDate(f!txtTo), Nz(f!cboPromotion, ""), Nz(f!cboObservation, ""), traits, Nz(f!txtSource, ""))
    DoCmd.Close acForm, f.Name: UXHistoryLoad: UXPriorSave = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXDirectoryBatchOpen() As Boolean
    ClearDirectoryReportPreview
    DoCmd.OpenForm "frmUXDirectoryBatch"
    Forms!frmUXDirectoryBatch!cmdApply.Enabled = False: Forms!frmUXDirectoryBatch!txtStamp = Null
    UXMemberRankChoices Forms!frmUXDirectoryBatch
    UXDirectoryRankFilter
    Forms!frmUXDirectoryBatch!cboGroup.RowSource = "SELECT GroupID,Title FROM SummaryGroup WHERE IsArchived=False ORDER BY Title"
    UXDirectoryBatchOpen = True
End Function
'Rank/rate filter choices: paygrades first, then exact ranks/rates, from the
'member directory or (source "Report") the saved reports.
Public Sub UXMemberRankChoices(ByVal f As Access.Form, Optional ByVal source As String = "CommandMember")
    Dim db As DAO.Database, r As DAO.Recordset, grades As Object, grade As String, key As Variant
    Set db = CurrentDb: Set grades = CreateObject("Scripting.Dictionary")
    f!cboRankFilter.RowSourceType = "Value List": f!cboRankFilter.RowSource = "All ranks/rates;Unentered"
    Set r = db.OpenRecordset("SELECT DISTINCT RankRate FROM [" & source & "] WHERE IsArchived=False AND RankRate Is Not Null AND RankRate<>'' ORDER BY RankRate", dbOpenSnapshot)
    Do While Not r.EOF
        grade = ValidationWornPaygrade(CStr(r!RankRate))
        If grade <> "" Then grades(grade) = True
        r.MoveNext
    Loop
    For Each key In grades.Keys: f!cboRankFilter.AddItem CStr(key): Next
    If Not (r.BOF And r.EOF) Then r.MoveFirst
    Do While Not r.EOF
        If Not grades.Exists(CStr(r!RankRate)) Then f!cboRankFilter.AddItem """" & Replace(CStr(r!RankRate), """", """""") & """"
        r.MoveNext
    Loop
    r.Close: f!cboRankFilter = "All ranks/rates"
End Sub
Public Function UXMemberRankWhere(ByVal rank As String) As String
    If rank = "All ranks/rates" Or rank = "" Then Exit Function
    If rank = "Unentered" Then UXMemberRankWhere = " AND Len(Trim(Nz(RankRate,'')))=0": Exit Function
    UXMemberRankWhere = " AND (RankRate=" & QuoteSQL(rank) & " OR ValidationWornPaygrade(Nz(RankRate,''))=" & QuoteSQL(rank) & ")"
End Function
Public Function UXDirectoryRankFilter() As Boolean
    ClearDirectoryReportPreview
    With Forms!frmUXDirectoryBatch
        !lstMembers.RowSource = "SELECT MemberID,FullName,RankRate FROM CommandMember WHERE IsArchived=False" & UXMemberRankWhere(Nz(!cboRankFilter, "All ranks/rates")) & " ORDER BY FullName"
        !lstMembers.Requery: !cmdApply.Enabled = False: !txtStamp = Null: !txtPreview = Null
    End With
    UXDirectoryRankFilter = True
End Function
Private Function UXDirectoryBatchSelection() As String
    Dim row As Variant, result As String
    For Each row In Forms!frmUXDirectoryBatch!lstMembers.ItemsSelected
        If result <> "" Then result = result & ","
        result = result & Forms!frmUXDirectoryBatch!lstMembers.ItemData(row)
    Next
    UXDirectoryBatchSelection = result
End Function
Public Function UXDirectoryBatchPreview() As Boolean
    Dim f As Access.Form
    ClearDirectoryReportPreview
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryBatch: f!cmdApply.Enabled = False: f!txtStamp = Null
    If IsNull(f!cboGroup) Or Not IsDate(f!txtFrom) Or Not IsDate(f!txtTo) Then Err.Raise vbObjectError + 6992, "modUXDirectory.UXDirectoryBatchPreview", "Choose the group and both report dates."
    f!txtPreview = PreviewDirectoryReports(CLng(f!cboGroup), UXDirectoryBatchSelection(), CDate(f!txtFrom), CDate(f!txtTo))
    f!txtStamp = CStr(f!cboGroup) & "|" & CStr(CDate(f!txtFrom)) & "|" & CStr(CDate(f!txtTo)) & "|" & UXDirectoryBatchSelection()
    f!cmdApply.Enabled = True: UXDirectoryBatchPreview = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXDirectoryBatchApply() As Boolean
    Dim f As Access.Form, stamp As String, count As Long, outcome As String, message As String
    outcome = "Not started"
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryBatch
    stamp = CStr(Nz(f!cboGroup, 0)) & "|" & CStr(CDate(f!txtFrom)) & "|" & CStr(CDate(f!txtTo)) & "|" & UXDirectoryBatchSelection()
    If stamp <> Nz(f!txtStamp, "") Then Err.Raise vbObjectError + 6992, "modUXDirectory.UXDirectoryBatchApply", "Selection or dates changed. Preview again."
    count = ApplyDirectoryReports(outcome): f!cmdApply.Enabled = False: f!txtStamp = Null
    f!lblStatus.Caption = "Created " & count & " new drafts. Open the summary group to edit them.": UXDirectoryBatchApply = True: Exit Function
Failed:
    message = Err.Description: ClearDirectoryReportPreview
    f!cmdApply.Enabled = False: f!txtStamp = Null
    Select Case outcome
        Case "Committed": message = "Batch committed; the display update failed. Do not create it again. Reopen the summary group to inspect the saved drafts. " & message
        Case "Rolled back": message = "Batch rolled back. Preview again before retrying. " & message
        Case "Not started": message = "Batch did not start. Preview again before Apply. " & message
        Case Else: message = "Batch outcome could not be confirmed. Inspect saved drafts and history before retrying. " & message
    End Select
    f!lblStatus.Caption = message
End Function
Public Function UXHistoryConfirmIdentity() As Boolean
    Dim id As Long, f As Access.Form
    If IsNull(Forms!frmUXMemberHistory!lstHistory) Then Exit Function
    id = CLng(Forms!frmUXMemberHistory!lstHistory)
    If id >= 0 Then Forms!frmUXMemberHistory!lblStatus.Caption = "Choose a saved Swift report to confirm its member link.": Exit Function
    DoCmd.OpenForm "frmUXConfirmIdentity": Set f = Forms!frmUXConfirmIdentity
    f!txtReport = -id: f!txtRevision = CurrentRevision(-id): f!txtMember = historyMember
    f!cboMember.RowSource = "SELECT MemberID,FullName,RankRate FROM CommandMember WHERE IsArchived=False ORDER BY FullName"
    f!cboMember = historyMember: UXHistoryIdentityPreview
    UXHistoryConfirmIdentity = True
End Function
Public Function UXHistoryIdentityPreview() As Boolean
    Dim f As Access.Form, id As Long, memberID As Long
    On Error GoTo Failed
    Set f = Forms!frmUXConfirmIdentity: f!cmdApply.Enabled = False
    If IsNull(f!cboMember) Then Err.Raise vbObjectError + 6993, "modUXDirectory.UXHistoryIdentityPreview", "Choose an active directory member."
    id = CLng(f!txtReport): memberID = CLng(f!cboMember)
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 6993, "modUXDirectory.UXHistoryIdentityPreview", "Choose an active directory member."
    f!txtMember = memberID: f!txtVersion = MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & memberID)
    f!txtPreview = "Saved report: " & UXIdentity(MajorScalar("SELECT OfficialName FROM Report WHERE ReportID=" & id), MajorScalar("SELECT MemberName FROM Report WHERE ReportID=" & id)) & vbCrLf & "Current identity state: " & MajorScalar("SELECT MatchState FROM MemberReport WHERE ReportID=" & id) & vbCrLf & "Confirm directory member: " & MajorScalar("SELECT FullName FROM CommandMember WHERE MemberID=" & memberID) & " / " & Nz(MajorScalar("SELECT RankRate FROM CommandMember WHERE MemberID=" & memberID), "") & vbCrLf & "Report names, SSNs and saved revisions stay unchanged. Confirm only after checking that this is the same person."
    f!lblStatus.Caption = "": f!cmdApply.Enabled = True: UXHistoryIdentityPreview = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXHistoryApplyIdentity() As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXConfirmIdentity
    If Nz(f!cboMember, 0) <> Nz(f!txtMember, 0) Then Err.Raise vbObjectError + 6993, "modUXDirectory.UXHistoryApplyIdentity", "Member selection changed. Preview the identity link again."
    LinkReportDirectoryMember CLng(f!txtReport), CLng(f!txtMember), CLng(f!txtRevision), CLng(f!txtVersion), Nz(f!txtReason, "")
    DoCmd.Close acForm, f.Name: UXHistoryLoad: UXHistoryApplyIdentity = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
'Counseling list: id;date;outcome;counselor or reason;worksheet. The status line
'shows the schedule state for the member.
Public Function UXHistoryCounselingLoad() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, rows As String, s As Object
    Set f = Forms!frmUXMemberHistory
    Call UXAssertView(f.Name)
    If Not UXHasControl(f, "lstCounseling") Then Exit Function
    f!lstCounseling.RowSourceType = "Value List"
    Set r = CurrentDb.OpenRecordset("SELECT CounselingID,CounseledOn,Outcome,CounselorName,NotPerformedReason,WorksheetDisposition FROM MemberCounseling WHERE MemberID=" & historyMember & " AND IsArchived=" & IIf(Nz(f!chkArchivedCounseling, False), "True", "False") & " ORDER BY CounseledOn DESC", dbOpenSnapshot)
    Do While Not r.EOF
        If rows <> "" Then rows = rows & ";"
        rows = rows & r!CounselingID & ";" & Format$(r!CounseledOn, "yyyy-mm-dd") & ";" & r!Outcome & ";""" & Replace(Nz(r!CounselorName, r!NotPerformedReason), """", "'") & """;" & Nz(r!WorksheetDisposition, "")
        r.MoveNext
    Loop
    r.Close
    f!lstCounseling.RowSource = rows
    f!cmdCounselingEdit.Enabled = Not Nz(f!chkArchivedCounseling, False)
    f!cmdCounselingArchive.Caption = IIf(Nz(f!chkArchivedCounseling, False), "Restore selected counseling", "Archive selected counseling")
    Set s = CounselingMemberStatus(historyMember)
    f!lblCounselingStatus.Caption = "Counseling: " & s("Status") & IIf(IsNull(s("DueMonth")), "", " (" & Format$(Nz(s("DueMonth"), Date), "mmmm yyyy") & ")") & IIf(s("Note") = "", "", ". " & s("Note"))
    UXHistoryCounselingLoad = True
End Function
Public Function UXCounselingOpen(Optional ByVal createNew As Boolean = True, Optional ByVal member As Long = 0) As Boolean
    Dim f As Access.Form, r As DAO.Recordset
    counselingID = 0: counselingVersion = 0
    If member > 0 Then historyMember = member
    counselingMember = historyMember
    If Not createNew Then
        If Nz(Forms!frmUXMemberHistory!chkArchivedCounseling, False) Then Forms!frmUXMemberHistory!lblStatus.Caption = "Restore the counseling record before editing.": Exit Function
        If IsNull(Forms!frmUXMemberHistory!lstCounseling) Then Forms!frmUXMemberHistory!lblStatus.Caption = "Select a counseling record to edit.": Exit Function
        counselingID = CLng(Forms!frmUXMemberHistory!lstCounseling)
    End If
    DoCmd.OpenForm "frmUXCounselingEditor": Set f = Forms!frmUXCounselingEditor
    f!lblMember.Caption = Nz(MajorScalar("SELECT FullName FROM CommandMember WHERE MemberID=" & counselingMember), "")
    f!cboOutcome.RowSourceType = "Value List": f!cboOutcome.RowSource = """" & COUNSELING_COMPLETED & """;""" & COUNSELING_NOT_PERFORMED & """"
    f!cboDisposition.RowSourceType = "Value List": f!cboDisposition.RowSource = """Given to member"";""Destroyed"""
    f!cboReason.RowSourceType = "Value List": f!cboReason.RowSource = CounselingReasonChoices()
    f!cboOutcome = COUNSELING_COMPLETED: f!txtDate = Date
    If counselingID > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberCounseling WHERE CounselingID=" & counselingID, dbOpenSnapshot)
        counselingVersion = r!VersionCounter: counselingMember = r!MemberID
        f!txtDate = r!CounseledOn: f!cboOutcome = r!Outcome: f!txtCounselor = r!CounselorName: f!cboReason = r!NotPerformedReason
        f!cboDisposition = r!WorksheetDisposition: r.Close
    End If
    Call UXCounselingOutcome: UXCounselingOpen = True
End Function
Public Function UXCounselingOutcome() As Boolean
    Dim f As Access.Form, completed As Boolean
    Set f = Forms!frmUXCounselingEditor: completed = (Nz(f!cboOutcome, "") = COUNSELING_COMPLETED)
    f!txtCounselor.Enabled = completed: f!cboReason.Enabled = Not completed
    f!lblStatus.Caption = IIf(completed, "The counselor is typed into block 31 of the next Regular report; the member signs the worksheet itself.", "NOT PERF goes in block 30 and this reason in block 31 of the next Regular report.")
    UXCounselingOutcome = True
End Function
Public Function UXCounselingSave() As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXCounselingEditor
    If Not IsDate(f!txtDate) Then Err.Raise vbObjectError + 9510, "modUXDirectory.UXCounselingSave", "Enter the counseling date."
    Call SaveCounseling(counselingID, counselingVersion, counselingMember, CDate(f!txtDate), Nz(f!cboOutcome, ""), Nz(f!txtCounselor, ""), Nz(f!cboReason, ""), Nz(f!cboDisposition, ""), Null)
    DoCmd.Close acForm, f.Name
    UXCounselingRefresh
    UXCounselingSave = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXCounselingArchive(Optional ByVal confirmed As Boolean = False) As Boolean
    Dim id As Long, version As Variant, restoring As Boolean
    On Error GoTo Failed
    If IsNull(Forms!frmUXMemberHistory!lstCounseling) Then Forms!frmUXMemberHistory!lblStatus.Caption = "Select a counseling record to archive.": Exit Function
    id = CLng(Forms!frmUXMemberHistory!lstCounseling)
    restoring = Nz(Forms!frmUXMemberHistory!chkArchivedCounseling, False)
    If Not confirmed Then
        If MsgBox(IIf(restoring, "Restore the selected counseling record? It will count toward the schedule and report suggestions again.", "Archive the selected counseling record? It is retained but no longer counts toward the schedule or report suggestions."), vbYesNo + vbQuestion + vbDefaultButton2, IIf(restoring, "Restore counseling record", "Archive counseling record")) <> vbYes Then Exit Function
    End If
    version = MajorScalar("SELECT VersionCounter FROM MemberCounseling WHERE CounselingID=" & id)
    If restoring Then UXCounselingArchive = RestoreCounseling(id, CLng(Nz(version, 0))) Else UXCounselingArchive = ArchiveCounseling(id, CLng(Nz(version, 0)))
    UXCounselingRefresh
    Exit Function
Failed:
    Forms!frmUXMemberHistory!lblStatus.Caption = Err.Description
End Function
Private Sub UXCounselingRefresh()
    If CurrentProject.AllForms("frmUXMemberHistory").IsLoaded Then UXHistoryCounselingLoad
    If CurrentProject.AllForms("frmMyWork").IsLoaded Then UXMyWorkCounseling
End Sub
'Open the history of the member selected in My Work, so its record can be added.
Public Function UXDirectoryHistoryFor(ByVal member As Long) As Boolean
    historyMember = member
    DoCmd.OpenForm "frmUXMemberHistory": UXHistoryLoad: UXDirectoryHistoryFor = True
End Function
