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
    UXDirectoryMergePairs
    UXDirectoryMergeChanged
    UXDirectoryMergeLoad = True
End Function
'The directory-wide list of possible duplicates, strongest first.
Public Function UXDirectoryMergePairs() As Boolean
    Const MAX_ROWS As Long = 200
    Dim f As Access.Form, pairs As Collection, item As Object, shown As Long, dismissed As Long, label As String
    Set f = Forms!frmUXDirectoryMerge
    Set pairs = DirectoryDuplicatePairs(Nz(f!chkShowDismissed, False))
    f!lstPairs.RowSource = ""
    f!lstPairs.AddItem ";Match;Sailor;Possible duplicate;Why"
    For Each item In pairs
        If item("DismissalID") > 0 Then dismissed = dismissed + 1
        If shown < MAX_ROWS Then
            Select Case item("Strength")
                Case 1: label = "Likely"
                Case 2: label = "Probable"
                Case Else: label = "Possible"
            End Select
            If item("DismissalID") > 0 Then label = "Different"
            f!lstPairs.AddItem item("LowID") & "-" & item("HighID") & "-" & item("DismissalID") & ";" & label & ";" & MergeListText(MergePairName(item("LowID"))) & ";" & MergeListText(MergePairName(item("HighID"))) & ";" & MergeListText(item("Why"))
            shown = shown + 1
        End If
    Next
    label = "Possible duplicate Sailors: " & (pairs.Count - dismissed)
    If dismissed > 0 Then label = label & "; marked different: " & dismissed
    If pairs.Count > MAX_ROWS Then label = label & " (first " & MAX_ROWS & " shown)"
    f!lblPairs.Caption = label
    f!cmdDismiss.Enabled = False: f!cmdRestore.Enabled = False
    UXDirectoryMergePairs = True
End Function
Private Function MergePairName(ByVal id As Long) As String
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT FullName,RankRate FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
    If Not r.EOF Then MergePairName = r!FullName & " (" & IIf(Len(Nz(r!RankRate, "")) > 0, r!RankRate & ", ", "") & "ID " & id & ")"
    r.Close
End Function
Private Function MergeListText(ByVal value As String) As String
    MergeListText = """" & Replace(Replace(Replace(value, """", """"""), vbCr, " "), vbLf, " ") & """"
End Function
'Selecting a pair fills both members, with the suggested survivor on the right.
Public Function UXDirectoryMergePair() As Boolean
    Dim f As Access.Form, key As Variant, low As Long, high As Long, survivor As Long, dismissal As Long
    Set f = Forms!frmUXDirectoryMerge
    'Each row's hidden key is low-high-dismissal.
    key = Split(Nz(f!lstPairs, "") & "--", "-")
    low = Val(key(0)): high = Val(key(1)): dismissal = Val(key(2))
    If low = 0 Or high = 0 Then Exit Function
    survivor = DuplicateSurvivor(low, high)
    f!cboTarget = survivor: f!cboSource = IIf(survivor = low, high, low)
    UXDirectoryMergeChanged
    f!cmdDismiss.Enabled = (dismissal = 0): f!cmdRestore.Enabled = (dismissal > 0)
    If dismissal > 0 Then f!txtPreview = "This pair is marked as different Sailors. Choose List again to show it as a possible duplicate."
    UXDirectoryMergePair = True
End Function
Public Function UXDirectoryMergeDismiss() As Boolean
    Dim f As Access.Form
    Call UXAssertView("frmUXDirectoryMerge")
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryMerge
    DismissDuplicatePair Nz(f!cboSource, 0), Nz(f!cboTarget, 0), Nz(f!txtReason, "")
    f!cboSource = Null: f!cboTarget = Null: f!txtReason = Null
    Call UXDirectoryMergePairs: Call UXDirectoryMergeChanged
    f!txtPreview = "Marked as different Sailors. Neither profile changed. Tick Show pairs marked as different Sailors to list it again."
    UXDirectoryMergeDismiss = True: Exit Function
Failed:
    f!txtPreview = Err.Description
End Function
Public Function UXDirectoryMergeRestore() As Boolean
    Dim f As Access.Form, dismissal As Long
    Call UXAssertView("frmUXDirectoryMerge")
    On Error GoTo Failed
    Set f = Forms!frmUXDirectoryMerge
    dismissal = Val(Split(Nz(f!lstPairs, "") & "--", "-")(2))
    If dismissal = 0 Then Err.Raise 5, "modUXDirectoryMerge.UXDirectoryMergeRestore", "Select a pair marked as different Sailors."
    RestoreDuplicatePair dismissal
    f!cboSource = Null: f!cboTarget = Null
    Call UXDirectoryMergePairs: Call UXDirectoryMergeChanged
    f!txtPreview = "Listed again as a possible duplicate."
    UXDirectoryMergeRestore = True: Exit Function
Failed:
    f!txtPreview = Err.Description
End Function
Public Function UXDirectoryMergeChanged() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXDirectoryMerge: reviewedStamp = "": reviewedOwner = 0: f!cmdApply.Enabled = False
    f!txtPreview = "Select a pair above, or choose any two members (for a changed name), then preview."
    f!cmdDismiss.Enabled = (Nz(f!cboSource, 0) > 0 And Nz(f!cboTarget, 0) > 0 And Nz(f!cboSource, 0) <> Nz(f!cboTarget, 0))
    UXDirectoryMergeChanged = True
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
