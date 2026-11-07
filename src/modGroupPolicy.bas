Attribute VB_Name = "modGroupPolicy"
Option Explicit
Private memberGroup As Long
Private memberPreviewToken As String
Private Function PolicyDocument(ByVal groupID As Long) As Object
    Dim d As Object, raw As String
    Set d = CreateObject("MSXML2.DOMDocument.6.0"): d.async = False: d.resolveExternals = False
    d.setProperty "ProhibitDTD", True
    raw = Nz(DLookup("KeySpec", "SummaryGroup", "GroupID=" & groupID), "")
    If Len(raw) = 0 Then raw = "<SwiftGroupPolicy version=""1""/>"
    If Not d.LoadXML(raw) Then Err.Raise vbObjectError + 6600, "modGroupPolicy.PolicyDocument", "Existing group policy is not recognized; preserve it and inspect diagnostics."
    If d.documentElement.nodeName <> "SwiftGroupPolicy" Or Nz(d.documentElement.getAttribute("version"), "") <> "1" Then Err.Raise vbObjectError + 6600, "modGroupPolicy.PolicyDocument", "Existing group policy is not recognized; preserve it and inspect diagnostics."
    Set PolicyDocument = d
End Function
Private Sub PolicyPut(ByVal d As Object, ByVal key As String, ByVal value As String)
    Dim n As Object
    Set n = d.selectSingleNode("/SwiftGroupPolicy/" & key)
    If n Is Nothing Then Set n = d.createElement(key): d.documentElement.appendChild n
    n.Text = value
End Sub
Public Function GroupPolicyValue(ByVal groupID As Long, ByVal key As String) As String
    Dim d As Object, n As Object
    Select Case key
        Case "Workflow"
            GroupPolicyValue = Nz(DLookup("WorkflowState", "GroupLifecycle", "GroupID=" & groupID), ""): Exit Function
        Case "RSCAPostedBaseline", "RSCAGroupSum", "RSCAGroupCount", "RSCAReportIDs", "RSCACompletedAt"
            GroupPolicyValue = RSPostingValue(groupID, key): Exit Function
    End Select
    Set d = PolicyDocument(groupID): Set n = d.selectSingleNode("/SwiftGroupPolicy/" & key)
    If Not n Is Nothing Then GroupPolicyValue = n.Text
End Function
Public Function GroupWorkflow(ByVal groupID As Long) As String
    GroupWorkflow = GroupPolicyValue(groupID, "Workflow")
    If Len(GroupWorkflow) = 0 Then GroupWorkflow = "Initialized"
End Function
'Completing with a signed date marks every draft or approved member report signed
'in the same transaction; leaving Completed reopens the group's signed reports.
Public Sub SetGroupWorkflow(ByVal groupID As Long, ByVal expected As String, ByVal state As String, Optional ByVal signedOn As Variant = Null)
    Dim w As DAO.Workspace, db As DAO.Database, d As Object, old As String, op As Long, failure As Long, message As String, author As Long
    author = MajorAuthor()
    If state <> "Initialized" And state <> "In-Work" And state <> "Ready for Signature" And state <> "Completed" Then Err.Raise vbObjectError + 6810, "modGroupPolicy.SetGroupWorkflow", "Choose a supported group status."
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    If GroupStamp(groupID) <> expected Then Err.Raise vbObjectError + 6811, "modGroupPolicy.SetGroupWorkflow", "Group changed. Refresh and choose its status again."
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 6812, "modGroupPolicy.SetGroupWorkflow", "Recover this group before changing its status."
    old = GroupWorkflow(groupID)
    If old = state And state <> "Completed" Then w.CommitTrans: Exit Sub
    Dim completedAt As Variant, signoffs As String, recorded As String
    If old = "Completed" Then completedAt = DLookup("ChangedAt", "GroupLifecycle", "GroupID=" & groupID)
    Set d = PolicyDocument(groupID)
    If state = "Completed" Then RSPostGroupCompletion groupID, d
    db.Execute "DELETE FROM GroupLifecycle WHERE GroupID=" & groupID, dbFailOnError
    db.Execute "INSERT INTO GroupLifecycle (GroupID,WorkflowState,ChangedBy,ChangedAt) VALUES (" & groupID & "," & QuoteSQL(state) & "," & author & ",Now())", dbFailOnError
    Dim reportsChanged As Long, outcome As String
    outcome = "Local group status changed; member report states unchanged."
    'Completing again also signs any member left unsigned, so a Completed group
    'never posts a draft without its signed date.
    If state = "Completed" And Not IsNull(signedOn) Then
        reportsChanged = MarkGroupReportsSigned(groupID, CDate(signedOn), signoffs): outcome = "Group completed; " & reportsChanged & " member reports marked signed " & Format$(signedOn, "yyyy-mm-dd") & "."
        'Remember exactly which sign-offs the completion made; re-completion adds to them.
        recorded = GroupPolicyValue(groupID, "CompletionSignoffs")
        If Len(signoffs) > 0 Then recorded = recorded & IIf(Len(recorded) > 0, ",", "") & signoffs
        PolicyPut d, "CompletionSignoffs", recorded
        db.Execute "UPDATE SummaryGroup SET KeySpec=" & QuoteSQL(d.XML) & " WHERE GroupID=" & groupID, dbFailOnError
    ElseIf old = "Completed" And state <> "Completed" Then
        reportsChanged = ReopenGroupReports(groupID, "Group reopened from Completed to " & state, completedAt): outcome = "Group reopened; " & reportsChanged & " member reports signed by the completion reopened for correction; reports signed separately keep their signatures."
        Set d = PolicyDocument(groupID): PolicyPut d, "CompletionSignoffs", ""
        db.Execute "UPDATE SummaryGroup SET KeySpec=" & QuoteSQL(d.XML) & " WHERE GroupID=" & groupID, dbFailOnError
    End If
    db.Execute "UPDATE SummaryGroup SET VersionCounter=VersionCounter+1 WHERE GroupID=" & groupID, dbFailOnError
    op = BeginOperation("Change group workflow", "Group " & groupID & "; BEFORE: " & old & "; AFTER: " & state)
    FinishOperation op, "Completed", outcome
    w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modGroupPolicy", "SetGroupWorkflow", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
Public Function GroupReviewCurrent(ByVal groupID As Long) As Boolean
    GroupReviewCurrent = GroupPolicyValue(groupID, "Complete") = "Yes" And GroupPolicyValue(groupID, "RosterStamp") = GroupStamp(groupID)
    If GroupReviewCurrent And CurrentProject.AllForms("frmEditor").IsLoaded Then
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)) > 0 Then
            If EditorHasPending() Then GroupReviewCurrent = False
        End If
    End If
End Function
Public Function GroupReviewLabel(ByVal groupID As Long) As String
    If GroupReviewCurrent(groupID) Then
        GroupReviewLabel = "Roster completeness recorded by " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & Val(GroupPolicyValue(groupID, "ReviewedBy"))), "Unknown contributor") & " at " & GroupPolicyValue(groupID, "ReviewedAt") & ". Local confirmation; not Navy certification."
    ElseIf GroupPolicyValue(groupID, "Complete") = "Yes" Then
        GroupReviewLabel = "Roster review is stale. Recheck membership, saved reports and policy context."
    Else
        GroupReviewLabel = "Roster completeness has not been recorded. Capacity remains provisional."
    End If
End Function
Public Sub SaveGroupPolicy(ByVal groupID As Long, ByVal expectedStamp As String, ByVal uicMode As String, ByVal officerBasis As String, ByVal note As String, ByVal complete As Boolean)
    Dim d As Object, w As DAO.Workspace, db As DAO.Database, old As String, op As Long, author As Long, failure As Long, message As String, started As Boolean, kind As String
    author = MajorAuthor(): kind = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & groupID), "")
    If kind = "FITREP" Then
        If uicMode <> "Officer normal" And uicMode <> "Officer dual-hatted" Then Err.Raise vbObjectError + 6601, "modGroupPolicy.SaveGroupPolicy", "Choose an officer command grouping mode."
    Else
        If uicMode <> "Unconfirmed" And uicMode <> "Combined enlisted UICs" And uicMode <> "Separate enlisted UICs" Then Err.Raise vbObjectError + 6601, "modGroupPolicy.SaveGroupPolicy", "Choose an enlisted UIC grouping mode."
    End If
    If officerBasis <> "Unconfirmed" And officerBasis <> "Active" And officerBasis <> "Reserve" Then Err.Raise vbObjectError + 6602, "modGroupPolicy.SaveGroupPolicy", "Choose a competitive-category basis."
    If Len(Trim$(note)) = 0 Or Len(note) > 2000 Then Err.Raise vbObjectError + 6603, "modGroupPolicy.SaveGroupPolicy", "Record the review basis and any UIC authority in a note of up to 2000 characters."
    If complete And DCount("*", "GroupMember", "GroupID=" & groupID) = 0 Then Err.Raise vbObjectError + 6604, "modGroupPolicy.SaveGroupPolicy", "An empty roster cannot be confirmed complete."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)) > 0 Then
            If EditorHasPending() Then Err.Raise vbObjectError + 6605, "modGroupPolicy.SaveGroupPolicy", "Save or discard pending member edits before recording a roster review."
        End If
    End If
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans: started = True
    If GroupStamp(groupID) <> expectedStamp Then Err.Raise vbObjectError + 6606, "modGroupPolicy.SaveGroupPolicy", "Group changed. Reopen policy review."
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 6607, "modGroupPolicy.SaveGroupPolicy", "Recover this group first."
    old = Nz(DLookup("KeySpec", "SummaryGroup", "GroupID=" & groupID), ""): Set d = PolicyDocument(groupID)
    PolicyPut d, "UICMode", uicMode: PolicyPut d, "OfficerBasis", officerBasis: PolicyPut d, "Note", note
    PolicyPut d, "Complete", IIf(complete, "Yes", "No"): PolicyPut d, "ReviewedBy", CStr(author): PolicyPut d, "ReviewedAt", Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    db.Execute "UPDATE SummaryGroup SET VersionCounter=VersionCounter+1 WHERE GroupID=" & groupID, dbFailOnError
    PolicyPut d, "RosterStamp", GroupStamp(groupID)
    db.Execute "UPDATE SummaryGroup SET KeySpec=" & QuoteSQL(d.XML) & " WHERE GroupID=" & groupID, dbFailOnError
    op = BeginOperation("Review summary-group policy", "Group " & groupID & vbCrLf & "BEFORE: " & old & vbCrLf & "AFTER: " & d.XML)
    FinishOperation op, "Completed", "Policy context and deliberate local completeness review recorded; report values unchanged."
    w.CommitTrans: started = False: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modGroupPolicy", "SaveGroupPolicy", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If started Then SwiftRollback w, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
Public Function UXOpenGroupPolicy() As Boolean
    Dim id As Long
    id = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If id > 0 Then DoCmd.OpenForm "frmUXGroupPolicy", , , , , , CStr(id): UXOpenGroupPolicy = True
End Function
Public Function UXLoadGroupPolicy() As Boolean
    Dim id As Long, kind As String
    With Forms!frmUXGroupPolicy
        id = Val(Nz(.OpenArgs, "0")): !groupID = id: !baseline = GroupStamp(id)
        kind = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & id), "")
        !cboUIC.RowSource = IIf(kind = "FITREP", "Officer normal;Officer dual-hatted", "Unconfirmed;Combined enlisted UICs;Separate enlisted UICs")
        !cboUIC = GroupPolicyValue(id, "UICMode"): If Len(Nz(!cboUIC, "")) = 0 Then !cboUIC = IIf(kind = "FITREP", "Officer normal", "Unconfirmed")
        !cboBasis = GroupPolicyValue(id, "OfficerBasis"): If Len(Nz(!cboBasis, "")) = 0 Then !cboBasis = "Unconfirmed"
        !cboBasis.Enabled = (kind = "FITREP"): !txtNote = GroupPolicyValue(id, "Note"): !chkComplete = False
        !txtEvidence = GroupReviewLabel(id) & vbCrLf & ValidationPolicyReference(IIf(kind = "FITREP", "GROUP-OFFICER", "GROUP-ENLISTED")) & vbCrLf & ValidationPolicyReference("GROUP-CATEGORY")
    End With
    UXLoadGroupPolicy = True
End Function
Public Function UXSaveGroupPolicy() As Boolean
    On Error GoTo Failed
    With Forms!frmUXGroupPolicy
        SaveGroupPolicy CLng(!groupID), CStr(!baseline), Nz(!cboUIC, ""), Nz(!cboBasis, "Unconfirmed"), Nz(!txtNote, ""), Nz(!chkComplete, False)
    End With
    DoCmd.Close acForm, "frmUXGroupPolicy", acSaveNo
    UXSaveGroupPolicy = True: Exit Function
Failed:
    Forms!frmUXGroupPolicy!lblStatus.Caption = Err.Description
End Function
Public Function GroupImpactText(ByVal groupID As Long, Optional ByVal overrides As Object, Optional ByVal membersCSV As String = "@Current") As String
    Dim item As Object, findings As Collection, n As Long, db As DAO.Database, r As DAO.Recordset, ids As Object, key As Variant, trait As Variant, f As Object, observed As Long, marks As Long, total As Double
    Set ids = CreateObject("Scripting.Dictionary")
    If membersCSV = "@Current" Then
        Set db = CurrentDb: Set r = db.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID, dbOpenSnapshot)
        Do While Not r.EOF: ids(CStr(r!ReportID)) = r!ReportID: r.MoveNext: Loop
        r.Close
    ElseIf Len(membersCSV) > 0 Then
        Set ids = MajorIDs(membersCSV)
    End If
    For Each key In ids.Keys
        Set f = SavedFields(CLng(ids(key)))
        If Not overrides Is Nothing Then
            If overrides.Exists(CStr(ids(key))) Then Set f = overrides(CStr(ids(key)))
        ElseIf membersCSV = "@Current" And CurrentProject.AllForms("frmEditor").IsLoaded Then
            If Nz(Forms!frmEditor!txtReportID, 0) = CLng(ids(key)) Then Set f = EditorPendingFields()
        End If
        If GroupObservedRecommendation(f) Then
            observed = observed + 1
            For Each trait In TraitKeys()
                If IsNumeric(Nz(f(trait), "")) Then
                    If CDbl(f(trait)) >= 1 And CDbl(f(trait)) <= 5 Then total = total + CDbl(f(trait)): marks = marks + 1
                End If
            Next
        End If
    Next
    Set findings = ValidateLiveGroup(groupID, overrides, membersCSV)
    GroupImpactText = ids.Count & " reports; " & observed & " observed; " & marks & " graded observed traits; arithmetic average: " & IIf(marks = 0, "Unavailable", Format$(total / IIf(marks = 0, 1, marks), "0.00")) & vbCrLf & GroupReviewLabel(groupID)
    If observed = 0 Then GroupImpactText = GroupImpactText & vbCrLf & "No observed reports; promotion capacity is unavailable."
    For Each item In findings
        If item("Code") = "CAPACITY" Then GroupImpactText = GroupImpactText & vbCrLf & item("Message")
        If Left$(item("Code"), 6) = "GROUP-" Then n = n + 1
    Next
    GroupImpactText = GroupImpactText & vbCrLf & n & " grouping findings; " & findings.Count & " total report/group findings. Open group validation for current saved/pending details."
    If Not overrides Is Nothing Or membersCSV <> "@Current" Then GroupImpactText = GroupImpactText & vbCrLf & "Proposed changes require a fresh roster review after Apply. Preview creates no report revision."
End Function
Public Function GroupBulkImpact(ByVal groupID As Long, ByVal batchID As Long) As String
    Dim changes As Object, f As Object, r As DAO.Recordset, db As DAO.Database, id As String
    Set changes = CreateObject("Scripting.Dictionary"): Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT * FROM qryChangeItemRevisionIDs WHERE BatchID=" & batchID & " ORDER BY ReportID,ItemID", dbOpenSnapshot)
    Do While Not r.EOF
        id = CStr(r!ReportID)
        If Not changes.Exists(id) Then changes.Add id, SavedFields(CLng(id))
        Set f = changes(id)
        If r!NewIsNull Then f(r!FieldKey) = Null Else f(r!FieldKey) = r!NewValue
        r.MoveNext
    Loop
    r.Close
    GroupBulkImpact = "BEFORE GROUP:" & vbCrLf & GroupImpactText(groupID) & vbCrLf & vbCrLf & "AFTER PROPOSED GROUP:" & vbCrLf & GroupImpactText(groupID, changes)
End Function
'UI membership previews carry a saved candidate-revision baseline as well as selection.
Public Sub GroupMembersStart(ByVal groupID As Long)
    memberGroup = groupID: memberPreviewToken = ""
End Sub
Private Function SelectedMembers() As String
    Dim f As Access.Form, i As Long, id As Variant
    Set f = Forms!frmUXMembers
    For i = 0 To f!lstReports.ListCount - 1
        id = f!lstReports.Column(0, i)
        If IsNumeric(id) And f!lstReports.Selected(i) Then
            If Len(SelectedMembers) > 0 Then SelectedMembers = SelectedMembers & ","
            SelectedMembers = SelectedMembers & CStr(id)
        End If
    Next
End Function
Private Function MemberToken(ByVal csv As String) As String
    Dim ids As Object, key As Variant
    MemberToken = GroupStamp(memberGroup) & "|selection:" & csv & "|note:" & Nz(Forms!frmUXMembers!txtNote, "")
    If Len(csv) = 0 Then Exit Function
    Set ids = MajorIDs(csv)
    For Each key In ids.Keys
        MemberToken = MemberToken & "|" & key & ":" & CurrentRevision(CLng(ids(key))) & ":" & Nz(DLookup("IsArchived", "Report", "ReportID=" & ids(key)), True) & ":" & Nz(DLookup("WorkflowState", "Report", "ReportID=" & ids(key)), "")
    Next
End Function
Public Function UXPreviewGroupMembers() As Boolean
    Dim csv As String
    On Error GoTo Failed
    csv = SelectedMembers()
    Forms!frmUXMembers!txtImpact = "BEFORE GROUP:" & vbCrLf & GroupImpactText(memberGroup) & vbCrLf & vbCrLf & "AFTER SELECTED MEMBERS:" & vbCrLf & GroupImpactText(memberGroup, , csv)
    memberPreviewToken = MemberToken(csv): UXPreviewGroupMembers = True: Exit Function
Failed:
    memberPreviewToken = "": Forms!frmUXMembers!lblStatus.Caption = Err.Description
End Function
Public Function UXApplyGroupMembers() As Boolean
    On Error GoTo Failed
    If Len(memberPreviewToken) = 0 Or memberPreviewToken <> MemberToken(SelectedMembers()) Then Err.Raise vbObjectError + 6608, "modGroupPolicy.UXApplyGroupMembers", "Selection, candidates or group changed. Preview membership again."
    UXApplyGroupMembers = UXSaveMembers(): If UXApplyGroupMembers Then memberPreviewToken = ""
    Exit Function
Failed:
    Forms!frmUXMembers!lblStatus.Caption = Err.Description
End Function
