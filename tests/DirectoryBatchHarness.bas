Attribute VB_Name = "DirectoryBatchHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6993, , message
    n = n + 1
End Sub
Private Function ApplyRejected() As Boolean
    On Error GoTo Rejected
    Call ApplyDirectoryReports(): Exit Function
Rejected:
    ApplyRejected = True
End Function
Public Function TestDirectoryBatch() As String
    Dim m As Long, other As Long, g As Long, count As Long, id As Long, r As DAO.Recordset, text As String, kind As Variant, rate As String, grade As String, newMember As Long, stamp As String
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
