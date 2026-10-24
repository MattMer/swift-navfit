Attribute VB_Name = "modUXDirectoryMerge"
Option Explicit
Private reviewedStamp As String
Private reviewedOwner As Long
Public Function UXDirectoryMergeOpen() As Boolean
    DoCmd.OpenForm "frmUXDirectoryMerge"
    UXDirectoryMergeOpen = True
End Function
Public Function UXDirectoryMergeLoad() As Boolean
    Dim f As Access.Form, sql As String
    Call UXAssertView("frmUXDirectoryMerge")
    Set f = Forms!frmUXDirectoryMerge
    sql = "SELECT MemberID,FullName & ' [ID ' & MemberID & ']' FROM CommandMember WHERE IsArchived=False ORDER BY FullName,MemberID"
    f!cboSource.RowSource = sql: f!cboTarget.RowSource = sql
    f!lstHistory.RowSource = "SELECT MergeID,FromMemberID AS [Original ID],ToMemberID AS [Survivor ID],MergeState AS [State],CreatedAt AS [Merged at] FROM DirectoryMerge ORDER BY MergeID DESC"
    UXDirectoryMergeChanged
    UXDirectoryMergeLoad = True
End Function
Public Function UXDirectoryMergeChanged() As Boolean
    Dim f As Access.Form, id As Long, surname As String
    Set f = Forms!frmUXDirectoryMerge: reviewedStamp = "": reviewedOwner = 0: f!cmdApply.Enabled = False
    f!txtPreview = "Choose the original member and the surviving member, then preview. Matching surnames suggest candidates only. For a changed name, choose any active member in the survivor list."
    id = Nz(f!cboSource, 0)
    surname = DirectorySurname(DLookup("FullName", "CommandMember", "MemberID=" & id), DLookup("ReportName", "CommandMember", "MemberID=" & id))
    f!lstCandidates.RowSource = "SELECT MemberID,FullName AS [Possible duplicate],RankRate AS [Rank/rate] FROM CommandMember WHERE IsArchived=False AND MemberID<>" & id & " AND DirectorySurname(FullName,ReportName)=" & QuoteSQL(surname) & " AND " & IIf(Len(surname) > 0, "True", "False") & " ORDER BY FullName"
    UXDirectoryMergeChanged = True
End Function
Public Function UXDirectoryMergeCandidate() As Boolean
    If IsNull(Forms!frmUXDirectoryMerge!lstCandidates) Then Exit Function
    Forms!frmUXDirectoryMerge!cboTarget = Forms!frmUXDirectoryMerge!lstCandidates
    UXDirectoryMergeCandidate = UXDirectoryMergeChanged()
End Function
Private Function MergeMemberSummary(ByVal id As Long, ByVal heading As String) As String
    Dim db As DAO.Database, r As DAO.Recordset, text As String, ssn As String
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
    ssn = Nz(r!MemberSSN, ""): If Len(ssn) > 0 Then ssn = "***-**-" & Right$(ssn, 4) Else ssn = "Unentered"
    text = heading & ": " & r!FullName & " [ID " & id & "]" & vbCrLf & "Report name: " & Nz(r!ReportName, "") & " | Rank: " & Nz(r!RankRate, "") & " | SSN: " & ssn & " | Source: " & Nz(r!SourceKind, "") & vbCrLf: r.Close
    Set r = db.OpenRecordset("SELECT r.ReportID,r.MemberName,r.PeriodEnd,r.MemberDoDID,m.MatchState FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & id & " ORDER BY r.ReportID", dbOpenSnapshot)
    text = text & "Linked reports (including archived reports):" & vbCrLf
    Do While Not r.EOF
        ssn = Nz(r!MemberDoDID, ""): If Len(ssn) > 0 Then ssn = "******" & Right$(ssn, 4) Else ssn = "Unentered"
        text = text & "  Report " & r!ReportID & ": " & Nz(r!MemberName, "") & " | " & Nz(r!PeriodEnd, "") & " | " & Nz(r!MatchState, "") & " | DoD ID " & ssn & vbCrLf: r.MoveNext
    Loop
    r.Close: Set r = db.OpenRecordset("SELECT PriorID,FormKind,PeriodEnd FROM PriorEvaluation WHERE MemberID=" & id & " ORDER BY PriorID", dbOpenSnapshot)
    text = text & "Prior evaluations (including archived entries):" & vbCrLf
    Do While Not r.EOF
        text = text & "  Prior " & r!PriorID & ": " & r!FormKind & " ending " & r!PeriodEnd & vbCrLf: r.MoveNext
    Loop
    r.Close: MergeMemberSummary = text
End Function
Public Function UXDirectoryMergePreview() As Boolean
    Dim f As Access.Form, a As Long, b As Long
    Call UXAssertView("frmUXDirectoryMerge")
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryMerge: reviewedStamp = "": f!cmdApply.Enabled = False
    reviewedOwner = 0: a = Nz(f!cboSource, 0): b = Nz(f!cboTarget, 0)
    reviewedStamp = PreviewDirectoryMerge(a, b)
    reviewedOwner = ActiveContributorID()
    f!txtPreview = MergeMemberSummary(a, "ORIGINAL - will be archived") & vbCrLf & MergeMemberSummary(b, "SURVIVOR - keeps profile") & vbCrLf & "All original report and prior links move to the survivor. Profiles are retained; profile fields are not combined. Report values, revisions, notes, review history and prior traits stay unchanged. Provisional/unconfirmed link states stay unchanged. Undo is available only while the affected data remains unchanged. Check that these are the same Sailor, then enter a reason and apply."
    f!cmdApply.Enabled = True: UXDirectoryMergePreview = True: Exit Function
Failed:
    reviewedStamp = "": f!txtPreview = Err.Description
End Function
Public Function UXDirectoryMergeApply() As Boolean
    Dim f As Access.Form, id As Long, message As String
    Call UXAssertView("frmUXDirectoryMerge")
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryMerge
    If reviewedOwner = 0 Or reviewedStamp = "" Then Err.Raise 5, "modUXDirectoryMerge.UXDirectoryMergeApply", "Preview the selected members first."
    If reviewedOwner <> ActiveContributorID() Then Err.Raise 5, "modUXDirectoryMerge.UXDirectoryMergeApply", "Preview the selected members as the current contributor before Apply."
    id = ApplyDirectoryMerge(Nz(f!cboSource, 0), Nz(f!cboTarget, 0), reviewedStamp, Nz(f!txtReason, ""))
    reviewedStamp = "": reviewedOwner = 0
    f!cboSource = Null: f!cboTarget = Null: f!txtReason = Null: UXDirectoryMergeLoad
    f!txtPreview = "Merge " & id & " completed. Select it in merge history to undo, provided no later changes affect these members."
    If CurrentProject.AllForms("frmUXDirectory").IsLoaded Then UXDirectoryLoad
    UXDirectoryMergeApply = True: Exit Function
Failed:
    message = Err.Description: reviewedStamp = "": reviewedOwner = 0
    f!cmdApply.Enabled = False
    If id > 0 Then message = "Merge " & id & " committed; the display update failed. Reopen the directory and inspect merge history before making further changes. " & message
    f!txtPreview = message
End Function
Public Function UXDirectoryMergeUndo() As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryMerge
    If Nz(f!lstHistory, 0) = 0 Then Err.Raise 5, "modUXDirectoryMerge.UXDirectoryMergeUndo", "Select a merge in the history list."
    UndoDirectoryMerge CLng(f!lstHistory): UXDirectoryMergeLoad
    f!txtPreview = "Undo completed. Original links and profile availability restored; merge history retained."
    If CurrentProject.AllForms("frmUXDirectory").IsLoaded Then UXDirectoryLoad
    UXDirectoryMergeUndo = True: Exit Function
Failed:
    f!txtPreview = Err.Description
End Function
