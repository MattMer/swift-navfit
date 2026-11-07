Attribute VB_Name = "modReportDirectory"
Option Compare Binary
Option Explicit
Private sourceReport As Long, sourceGroup As Long, sourceValues As Object, sourceSeniorKey As String

'Only fields actually held by the member record; unentered values never clear reports.
Public Function MemberReportValues(ByVal memberID As Long) As Object
    Dim r As DAO.Recordset, values As Object, key As Variant, ruad As Object
    Set values = CreateObject("Scripting.Dictionary")
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE IsArchived=False AND MemberID=" & memberID, dbOpenSnapshot)
    If Not r.EOF Then
        values("MemberName") = r!FullName.Value: values("OfficialName") = r!ReportName.Value
        values("RankRate") = r!RankRate.Value: values("MemberSSN") = r!MemberSSN.Value
        If Not IsNull(r!DateReported) Then values("DateReported") = Format$(r!DateReported, "yyyy-mm-dd")
        Set ruad = RUADDecode(r!RUADData.Value)
        values("UnitUIC") = ruad("UIC"): values("Designator") = ruad("DESG")
    End If
    r.Close
    For Each key In values.Keys
        If Len(Trim$(Nz(values(key), ""))) = 0 Then values.Remove CStr(key)
    Next
    Set MemberReportValues = values
End Function

'A saved member profile supplies its open reports: confirmed links and drafts only.
'Signed, locally approved, under-review reports, a report open in the editor and a
'rank/rate that would leave a summary group's paygrade are skipped and named.
'The caller owns the transaction. Returns a summary for the contributor.
Public Function UpdateMemberOpenReports(ByVal memberID As Long) As String
    Dim r As DAO.Recordset, values As Object, saved As Object, changes As Object, key As Variant
    Dim updated As Long, skipped As String, reason As String, holderWhere As String, editorID As Long, otherGrade As Variant
    Set values = MemberReportValues(memberID)
    If values.Count = 0 Then Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then editorID = Nz(Forms!frmEditor!txtReportID, 0)
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID,r.CurrentRevisionID,r.WorkflowState,r.FormKind FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND m.MatchState='Confirmed' AND r.IsArchived=False ORDER BY r.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        Set saved = SavedFields(r!ReportID): Set changes = CreateObject("Scripting.Dictionary")
        For Each key In values.Keys
            If UXFieldApplies(CStr(key), CStr(r!FormKind)) Then
                If Not SameFieldValue(saved(key), values(key)) Then changes(CStr(key)) = values(key)
            End If
        Next
        If changes.Count > 0 Then
            reason = ""
            If r!WorkflowState = "Signed" Then
                reason = "signed"
            ElseIf r!WorkflowState <> "Draft" Then
                reason = "locally approved"
            ElseIf r!ReportID = editorID Then
                reason = "open in the editor"
            ElseIf ReviewHolder(r!ReportID, holderWhere) > 0 Then
                reason = "under review"
            ElseIf changes.Exists("RankRate") Then
                otherGrade = MajorScalar("SELECT TOP 1 m.Paygrade FROM GroupMember AS m INNER JOIN SummaryGroup AS g ON m.GroupID=g.GroupID WHERE m.ReportID=" & r!ReportID & " AND g.IsArchived=False AND m.Paygrade<>" & QuoteSQL(ValidationWornPaygrade(CStr(changes("RankRate")))))
                If Not IsNull(otherGrade) Then reason = "in a " & otherGrade & " summary group"
            End If
            If Len(reason) > 0 Then
                skipped = skipped & IIf(Len(skipped) > 0, "; ", "") & "report " & r!ReportID & " (" & reason & ")"
            Else
                Call CommitSavedFields(CLng(r!ReportID), CLng(r!CurrentRevisionID), changes, "Update from member profile", ownTransaction:=False, syncDirectory:=False)
                updated = updated + 1
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    If updated > 0 Then UpdateMemberOpenReports = "Updated " & updated & " open report" & IIf(updated = 1, "", "s") & " from this profile."
    If Len(skipped) > 0 Then UpdateMemberOpenReports = Trim$(UpdateMemberOpenReports & " Not updated: " & skipped & ".")
End Function

'A group's reporting senior comes from its selected senior record, not from
'defaults: new reports created in the group start with that record's fields.
Public Sub GroupSeniorCreationValues(ByVal groupID As Long, ByVal values As Object)
    Dim senior As Long, fields As Object, key As Variant
    If groupID <= 0 Then Exit Sub
    senior = Nz(MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & groupID), 0)
    If senior = 0 Then Exit Sub
    Set fields = RSSeniorReportValues(senior)
    For Each key In fields.Keys: values(CStr(key)) = fields(key): Next
End Sub

Public Function ReportDirectoryValues(ByVal reportID As Long, ByVal groupID As Long) As Object
    Dim fields As Object, values As Object, key As Variant, member As Long, senior As Long
    Set fields = CreateObject("Scripting.Dictionary")
    member = Nz(MajorScalar("SELECT l.MemberID FROM MemberReport AS l INNER JOIN CommandMember AS m ON l.MemberID=m.MemberID WHERE l.ReportID=" & reportID & " AND l.MatchState='Confirmed' AND m.IsArchived=False"), 0)
    If member > 0 Then
        Set values = MemberReportValues(member)
        For Each key In values.Keys: fields(CStr(key)) = values(key): Next
    End If
    If groupID > 0 Then senior = Nz(MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & groupID), 0)
    If senior > 0 Then
        Set values = RSSeniorReportValues(senior)
        For Each key In values.Keys
            'The selected senior is one complete record: do not retain old
            'name parts/contact fields when that record explicitly has no value.
            fields(CStr(key)) = values(key)
        Next
    End If
    Set ReportDirectoryValues = fields
End Function

'The signature covers report, group and source versions before confirmation and writes.
Public Function GroupDirectoryPlan(ByVal groupID As Long) As Object
    Dim plan As Object, items As New Collection, r As DAO.Recordset, fields As Object, saved As Object, changes As Object
    Dim key As Variant, signature As String, sourceSignature As String, member As Long, senior As Long, signed As Long, unconfirmed As Long, grade As String, memberStamp As String
    If groupID <= 0 Or Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 9870, "modReportDirectory", "Choose an active summary group."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 9870, "modReportDirectory", "Close the report editor before updating group data."
    signature = GroupStamp(groupID): grade = CStr(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupID))
    senior = Nz(MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & groupID), 0)
    signature = signature & "|S:" & senior
    If senior > 0 Then signature = signature & ":" & Nz(MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & senior), 0)
    sourceSignature = GroupRouteStamp(groupID) & "|G:" & Nz(MajorScalar("SELECT VersionCounter FROM SummaryGroup WHERE GroupID=" & groupID), 0) & ":" & grade & "|" & Mid$(signature, InStrRev(signature, "|S:") + 1)
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID,r.CurrentRevisionID,r.WorkflowState,r.FormKind FROM Report AS r INNER JOIN GroupMember AS g ON r.ReportID=g.ReportID WHERE g.GroupID=" & groupID & " AND r.IsArchived=False ORDER BY r.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        signature = signature & "|R:" & r!ReportID & ":" & r!CurrentRevisionID & ":" & r!WorkflowState
        member = Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & r!ReportID & " AND MatchState='Confirmed'"), 0)
        memberStamp = ":M:" & member
        If member > 0 Then
            memberStamp = memberStamp & ":" & Nz(MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & member), 0) & ":" & Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & member), True)
        Else
            unconfirmed = unconfirmed + 1
        End If
        signature = signature & memberStamp
        sourceSignature = sourceSignature & "|R:" & r!ReportID & ":" & r!FormKind & memberStamp
        Set saved = SavedFields(r!ReportID): Set fields = ReportDirectoryValues(r!ReportID, groupID)
        Set changes = CreateObject("Scripting.Dictionary")
        For Each key In fields.Keys
            If UXFieldApplies(CStr(key), CStr(r!FormKind)) Then
                If Not SameFieldValue(saved(key), fields(key)) Then changes(CStr(key)) = fields(key)
            End If
        Next
        If changes.Count > 0 Then
            If r!WorkflowState = "Signed" Then
                signed = signed + 1
            Else
                If r!WorkflowState <> "Draft" And r!WorkflowState <> "Approved" Then r.Close: Err.Raise vbObjectError + 9871, "modReportDirectory", "A report is under review. Finish or reopen its review before updating group data."
                If changes.Exists("RankRate") Then
                    If ValidationWornPaygrade(CStr(changes("RankRate"))) <> grade Then r.Close: Err.Raise vbObjectError + 9872, "modReportDirectory", "A directory rank/rate no longer matches this group's paygrade. Review group membership first. No reports changed."
                End If
                items.Add Array(CLng(r!ReportID), CLng(r!CurrentRevisionID), CStr(r!WorkflowState), changes)
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    Set plan = CreateObject("Scripting.Dictionary")
    plan.Add "GroupID", groupID: plan.Add "Signature", signature: plan.Add "Items", items
    plan.Add "Signed", signed: plan.Add "Unconfirmed", unconfirmed
    plan.Add "Owner", ActiveContributorID()
    plan.Add "SourceSignature", sourceSignature
    Set GroupDirectoryPlan = plan
End Function

Public Function ApplyGroupDirectoryPlan(ByVal plan As Object) As Long
    Dim fresh As Object, after As Object, w As DAO.Workspace, item As Variant, operation As Long, trans As Boolean, failure As Object
    On Error GoTo Failed
    If CLng(plan("Owner")) <> ActiveContributorID() Then Err.Raise vbObjectError + 9873, "modReportDirectory", "The active contributor changed. Preview the update again."
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set fresh = GroupDirectoryPlan(CLng(plan("GroupID")))
    If fresh("Signature") <> plan("Signature") Then Err.Raise vbObjectError + 9873, "modReportDirectory", "A report, member, senior or group changed during confirmation. Preview the update again."
    If fresh("Items").Count > 0 Then
        operation = BeginOperation("Update group from directory", "Group " & plan("GroupID") & "; reports " & fresh("Items").Count)
        For Each item In fresh("Items")
            If item(2) = "Approved" Then ReopenApprovedReport CLng(item(0)), "Confirmed directory update"
            Call CommitSavedFields(CLng(item(0)), CLng(item(1)), item(3), "Update from linked member and reporting senior", ownTransaction:=False, syncDirectory:=False)
            AuditCorrelateRevision CurrentRevision(CLng(item(0))), operation
            ApplyGroupDirectoryPlan = ApplyGroupDirectoryPlan + 1
        Next
        Set after = GroupDirectoryPlan(CLng(plan("GroupID")))
        If after("SourceSignature") <> fresh("SourceSignature") Then Err.Raise vbObjectError + 9873, "modReportDirectory", "A linked record or group changed during the update. No reports changed; preview again."
        FinishOperation operation, "Completed", "Updated " & ApplyGroupDirectoryPlan & " reports; signed skipped " & fresh("Signed") & "; member identities unavailable " & fresh("Unconfirmed")
    End If
    w.CommitTrans: trans = False
    Call UXDirectorySourceInvalidate: UXEditorContextChanged
    Exit Function
Failed:
    Set failure = SwiftFailure("modReportDirectory", "ApplyGroupDirectoryPlan", "Apply linked records", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not started"))
    If trans Then SwiftRollback w, failure
    On Error GoTo 0: SwiftRaise failure
End Function

Public Function UXGroupDirectoryRefresh(ByVal screen As String, Optional ByVal decision As VbMsgBoxResult = 0) As Boolean
    Dim plan As Object, item As Variant, key As Variant, message As String, count As Long, id As Long
    On Error GoTo Failed
    If Not UXGridDrain() Then Exit Function
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        If UXGroupBulkPending() Then Err.Raise vbObjectError + 9870, "modReportDirectory", "Apply or discard pending shared-block edits first."
    End If
    id = UXResolvedGroupContext(screen): Set plan = GroupDirectoryPlan(id)
    message = "Update " & plan("Items").Count & " reports from their confirmed member records and the group's selected reporting senior?" & vbCrLf
    message = message & "Existing values in those fields will be replaced. Unentered member values are kept; selected senior fields include cleared values. Locally approved reports return to In-Work; signed reports are skipped." & vbCrLf
    For Each item In plan("Items")
        message = message & vbCrLf & "Report " & item(0) & ": "
        For Each key In item(3).Keys: message = message & UXFieldLabel(CStr(key)) & "; ": Next
    Next
    If plan("Items").Count > 0 Then
        If decision = 0 Then decision = MsgBox(Left$(message, 950), vbYesNo + vbQuestion + vbDefaultButton2, "Update from directory?")
        If decision <> vbYes Then Exit Function
    End If
    count = ApplyGroupDirectoryPlan(plan)
    If screen = "frmUXComparison" Then UXShowGroupMembers id, True
    DirectoryRefreshStatus screen, "Updated " & count & " reports. Signed skipped: " & plan("Signed") & ". Unconfirmed member links skipped: " & plan("Unconfirmed") & "."
    UXGroupDirectoryRefresh = True: Exit Function
Failed:
    DirectoryRefreshStatus screen, "Directory update: " & Err.Description
End Function

Public Function UXDirectorySourceInvalidate() As Boolean
    Set sourceValues = Nothing: sourceReport = 0: sourceGroup = 0: UXDirectorySourceInvalidate = True
End Function

'This is a current-value match indicator, not proof of import origin or a lock.
Public Sub UXDirectorySourceStyle()
    Dim f As Access.Form, c As Access.Control, id As Long, groupID As Long, key As String, linked As Boolean
    Dim seniorKey As String, senior As Long, values As Object, field As Variant
    If Not UXSurfaceAvailable() Then Exit Sub
    id = Nz(Forms!frmEditor!txtReportID, 0)
    groupID = Nz(Forms!frmEditor!uxReturnGroup, 0)
    If groupID = 0 Then groupID = Nz(Forms!frmEditor!uxGroup, 0)
    seniorKey = Nz(Forms!frmEditor!txtReportingSeniorSSN, "")
    If sourceValues Is Nothing Or id <> sourceReport Or groupID <> sourceGroup Or seniorKey <> sourceSeniorKey Then
        Set sourceValues = ReportDirectoryValues(id, groupID): sourceReport = id: sourceGroup = groupID
        sourceSeniorKey = seniorKey
        If Len(seniorKey) > 0 Then
            If DCount("*", "ReportingSeniorValue", "FieldKey='ReportingSeniorSSN' AND ValueText=" & QuoteSQL(seniorKey)) = 1 Then
                senior = CLng(MajorScalar("SELECT SeniorID FROM ReportingSeniorValue WHERE FieldKey='ReportingSeniorSSN' AND ValueText=" & QuoteSQL(seniorKey)))
                Set values = RSSeniorReportValues(senior)
                For Each field In values.Keys
                    sourceValues(CStr(field)) = values(field)
                Next
            End If
        End If
    End If
    Set f = Forms!frmEditor!uxFieldSurface.Form
    For Each c In f.Controls
        If Left$(c.Tag, 5) = "SURF:" And c.ControlType <> acCheckBox Then
            key = Mid$(c.Tag, 6): linked = False
            If sourceValues.Exists(key) Then
                If Len(Nz(sourceValues(key), "")) > 0 Then linked = SameFieldValue(Forms!frmEditor.Controls(c.Name).Value, sourceValues(key))
            End If
            If linked Then
                c.BackColor = UXColor("Inherited")
                c.ControlTipText = "Matches linked " & IIf(Left$(key, 15) = "ReportingSenior", "reporting senior", "member") & " record. Edit here to override."
            Else
                c.BackColor = UXColor("WhiteInput")
                If Left$(Nz(c.ControlTipText, ""), 14) = "Matches linked" Then c.ControlTipText = UXFieldCaption(key, Nz(Forms!frmEditor!cboKind, ""))
            End If
        End If
    Next
End Sub

Private Sub DirectoryRefreshStatus(ByVal screen As String, ByVal text As String)
    If screen = "frmUXBulk" Then
        Forms!frmUXBulk!uxStatus.Caption = text
    Else
        Forms(screen)!lblStatus.Caption = text
    End If
End Sub
