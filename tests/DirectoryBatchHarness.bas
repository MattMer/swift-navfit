Attribute VB_Name = "DirectoryBatchHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6993, , message
    n = n + 1
End Sub
Public Function PrepareDirectoryBatchState() As String
    Dim first As Long, second As Long, group As Long, i As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC DIRECTORY BATCH STATE", "Synthetic batch reviewer"
    first = SaveDirectoryMember(0, 0, "Synthetic First", "FIRST, SYNTHETIC", "IT2", "000000001")
    second = SaveDirectoryMember(0, 0, "Synthetic Second", "SECOND, SYNTHETIC", "IT2", "000000002")
    group = CreateSummaryGroup("Synthetic directory state group", "EVAL", "E5")
    If Not UXDirectoryBatchOpen() Then Err.Raise 5, , "Directory batch did not open"
    With Forms!frmUXDirectoryBatch
        !cboGroup = group: !txtFrom = #1/1/2026#: !txtTo = #9/30/2026#
        For i = 0 To !lstMembers.ListCount - 1
            If IsNumeric(!lstMembers.Column(0, i)) Then !lstMembers.Selected(i) = True
        Next
    End With
    If Not UXDirectoryBatchPreview() Then Err.Raise 5, , "Directory state preview failed"
    PrepareDirectoryBatchState = "PASS: directory batch state prepared": Exit Function
Failed:
    PrepareDirectoryBatchState = "FAIL: directory state preparation: " & Err.Number & " " & Err.Description
End Function
Public Function TestDirectoryBatchOutcome(ByVal fault As String) As String
    Dim prepared As String, expected As Long, message As String, context As Object
    On Error GoTo Failed
    n = 0
    prepared = PrepareDirectoryBatchState(): Check Left$(prepared, 5) = "PASS:", prepared
    If fault = "Rollback" Then CurrentDb.Execute "UPDATE CommandMember SET MemberSSN='bad' WHERE MemberID=" & DMax("MemberID", "CommandMember"), dbFailOnError
    Check Not UXDirectoryBatchApply(), "Injected directory batch failure was accepted"
    message = Forms!frmUXDirectoryBatch!lblStatus.Caption
    expected = IIf(fault = "Committed" Or fault = "Display", 2, 0)
    Check DCount("*", "Report") = expected And DCount("*", "GroupMember") = expected And DCount("*", "MemberReport") = expected, "Directory outcome disagrees with saved reports/links"
    Select Case fault
        Case "Rollback": Check InStr(message, "Batch rolled back") > 0, "Confirmed rollback was not disclosed"
        Case "Uncertain": Check InStr(message, "could not be confirmed") > 0 And InStr(message, "Nothing") = 0, "Uncertain outcome claims nothing committed"
        Case Else: Check InStr(message, "Batch committed") > 0 And InStr(message, "Nothing") = 0, "Committed outcome claims nothing committed"
    End Select
    If fault = "Committed" Then
        Set context = SwiftLastFailureContext()
        Check context("Number") = vbObjectError + 9692 And context("Transaction") = "Committed", "Postcommit directory primary identity/outcome lost"
    End If
    Check Not Forms!frmUXDirectoryBatch!cmdApply.Enabled And IsNull(Forms!frmUXDirectoryBatch!txtStamp), "Failed directory Apply retains action/stamp"
    Check Not UXDirectoryBatchApply(), "Directory Apply could repeat after failure"
    Check DCount("*", "Report") = expected And DCount("*", "GroupMember") = expected, "Repeated directory Apply duplicated data"
    TestDirectoryBatchOutcome = "PASS: " & n & " directory " & fault & " outcome/duplicate prevention assertions": Exit Function
Failed:
    TestDirectoryBatchOutcome = "FAIL: directory " & fault & ": " & Err.Number & " " & Err.Description
End Function
Public Function VerifyDirectoryBatchReset() As String
    On Error GoTo Failed
    n = 0
    Check Not UXDirectoryBatchApply(), "Directory batch preview survived reset"
    Check DCount("*", "Report") = 0 And DCount("*", "GroupMember") = 0, "Reset directory Apply created data"
    Check Not Forms!frmUXDirectoryBatch!cmdApply.Enabled And IsNull(Forms!frmUXDirectoryBatch!txtStamp), "Reset directory failure retained an actionable stamp"
    VerifyDirectoryBatchReset = "PASS: " & n & " real-reset directory preview/no-write assertions": Exit Function
Failed:
    VerifyDirectoryBatchReset = "FAIL: directory reset: " & Err.Number & " " & Err.Description
End Function
Private Function ApplyRejected() As Boolean
    On Error GoTo Rejected
    Call ApplyDirectoryReports(): Exit Function
Rejected:
    ApplyRejected = True
End Function
Private Function InvalidPreviewRejected(ByVal group As Long, ByVal member As Long) As Boolean
    Dim text As String
    On Error GoTo Rejected
    text = PreviewDirectoryReports(group, CStr(member), #9/30/2026#, #1/1/2026#)
    Exit Function
Rejected:
    InvalidPreviewRejected = True
End Function
Public Function TestDirectoryBatch() As String
    Dim m As Long, other As Long, g As Long, count As Long, id As Long, r As DAO.Recordset, text As String, kind As Variant, rate As String, grade As String, newMember As Long, stamp As String, owner As Long, second As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC DIRECTORY BATCH", "Tester"
    m = SaveDirectoryMember(0, 0, "Casey Morgan", "MORGAN, CASEY", "IT2", "000000005")
    other = SaveDirectoryMember(0, 0, "Jordan Reed", "REED, JORDAN", "IT2", "")
    g = CreateSummaryGroup("Directory group", "EVAL", "E5")
    SetFieldDefault "Command", 1, "ShipStation", "COMMAND"
    SetFieldDefault "Group", g, "ShipStation", "GROUP"
    text = PreviewDirectoryReports(g, m & "," & other, #1/1/2026#, #9/30/2026#)
    Check InStr(text, "Casey Morgan") > 0 And InStr(text, "Jordan Reed") > 0, "Preview missing member names"
    Check DCount("*", "Report") = 0, "Preview created reports"
    Check InvalidPreviewRejected(g, m), "Invalid dates were accepted"
    Check ApplyRejected() And DCount("*", "Report") = 0, "Failed preview retained the earlier actionable batch"
    text = PreviewDirectoryReports(g, m & "," & other, #1/1/2026#, #9/30/2026#)
    owner = ActiveContributorID(): second = AddContributor("Second synthetic batch creator")
    SwitchContributor second
    Check ApplyRejected() And DCount("*", "Report") = 0, "Another contributor applied the reviewed directory batch"
    SwitchContributor owner
    Check ApplyRejected(), "Rejected ownership retained an actionable directory batch"
    text = PreviewDirectoryReports(g, m & "," & other, #1/1/2026#, #9/30/2026#)
    Check ApplyDirectoryReports() = 2, "Batch count wrong"
    Check DCount("*", "GroupMember", "GroupID=" & g) = 2, "Group membership missing"
    Check DCount("*", "MemberReport", "MatchState='Confirmed'") = 2, "Chosen directory identities not confirmed"
    Check DCount("*", "Report", "ShipStation='GROUP'") = 2, "Group starting defaults not applied"
    Check DCount("*", "CommandMember") = 2, "Batch created duplicate members"
    Check ApplyRejected(), "Batch could be replayed"
    DoCmd.OpenForm "frmUXDirectory": UXDirectoryBatchOpen
    Forms!frmUXDirectoryBatch!cboGroup = g: Forms!frmUXDirectoryBatch!txtFrom = #1/1/2027#: Forms!frmUXDirectoryBatch!txtTo = #9/30/2027#
    Forms!frmUXDirectoryBatch!lstMembers.Selected(0) = True
    Check UXDirectoryBatchPreview(), "Native batch preview failed"
    Check UXDirectoryClose("frmUXDirectoryBatch"), "Batch close failed"
    Check ApplyRejected() And DCount("*", "Report") = 2, "Closed batch retained an actionable preview"
    Check UXDirectoryBatchOpen(), "Batch reopen failed"
    Forms!frmUXDirectoryBatch!cboGroup = g: Forms!frmUXDirectoryBatch!txtFrom = #1/1/2027#: Forms!frmUXDirectoryBatch!txtTo = #9/30/2027#
    Forms!frmUXDirectoryBatch!lstMembers.Selected(0) = True
    Check Not UXDirectoryBatchApply() And DCount("*", "Report") = 2, "Reopened batch reused its previous preview"
    Check UXDirectoryBatchPreview(), "Reopened batch could not preview again"
    Forms!frmUXDirectoryBatch!txtTo = #10/1/2027#
    Check Not UXDirectoryBatchApply() And DCount("*", "Report") = 2, "Changed UI period applied old preview"
    Check UXDirectoryBatchPreview(), "Native preview retry failed"
    Check UXDirectoryBatchApply() And DCount("*", "Report") = 3, "Native batch apply failed"
    text = PreviewDirectoryReports(g, CStr(other), #1/1/2028#, #9/30/2028#)
    SaveDirectoryMember other, 1, "Jordan Reed updated", "REED, JORDAN", "IT2", ""
    Check ApplyRejected() And DCount("*", "Report") = 3, "Stale profile batch was committed"
    text = PreviewDirectoryReports(g, m & "," & other, #1/1/2029#, #9/30/2029#)
    stamp = GroupStamp(g)
    CurrentDb.Execute "UPDATE CommandMember SET MemberSSN='bad' WHERE MemberID=" & other
    Check ApplyRejected(), "Late invalid identity did not reject batch"
    Check DCount("*", "Report") = 3 And DCount("*", "GroupMember") = 3 And GroupStamp(g) = stamp, "Late failure retained reports or memberships"
    CurrentDb.Execute "UPDATE CommandMember SET MemberSSN=Null WHERE MemberID=" & other
    text = PreviewDirectoryReports(g, CStr(other), #1/1/2030#, #9/30/2030#)
    SetFieldDefault "Command", 1, "UnitUIC", "12345"
    Check ApplyRejected() And DCount("*", "Report") = 3, "Changed global default did not reject batch"
    For Each kind In Array("CHIEFEVAL", "FITREP")
        If kind = "CHIEFEVAL" Then rate = "ITC": grade = "E7" Else rate = "LT": grade = "O3"
        newMember = SaveDirectoryMember(0, 0, "Synthetic " & kind, "SYNTHETIC, MEMBER", rate, "")
        g = CreateSummaryGroup("Synthetic " & kind, CStr(kind), grade)
        text = PreviewDirectoryReports(g, CStr(newMember), #1/1/2026#, #9/30/2026#)
        Check ApplyDirectoryReports() = 1, CStr(kind) & " creation failed"
        Check DCount("*", "Report", "FormKind='" & kind & "' AND RankRate='" & rate & "'") = 1, CStr(kind) & " official identity/type lost"
    Next
    UXDirectoryClose "frmUXDirectoryBatch": UXDirectoryClose "frmUXDirectory"
    TestDirectoryBatch = "PASS: " & n & " directory batch preview, atomic creation and stale-input assertions": Exit Function
Failed:
    TestDirectoryBatch = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
