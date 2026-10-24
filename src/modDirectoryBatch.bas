Attribute VB_Name = "modDirectoryBatch"
Option Explicit
Private selectedMembers As Object, memberStamps As Object, pendingGroup As Long
Private groupBaseline As String, defaultBaseline As String, fromDate As Date, toDate As Date
Public Function PreviewDirectoryReports(ByVal groupID As Long, ByVal memberCSV As String, ByVal startDate As Date, ByVal endDate As Date) As String
    Dim ids As Object, key As Variant, r As DAO.Recordset, kind As String, grade As String, result As String, field As Variant, defaultValue As Variant
    If endDate < startDate Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", "Report end precedes its start."
    Set selectedMembers = Nothing: Set memberStamps = CreateObject("Scripting.Dictionary")
    Set ids = MajorIDs(memberCSV)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM SummaryGroup WHERE GroupID=" & groupID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", "Choose an existing summary group."
    If r!IsArchived Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", "Recover the group first."
    kind = r!FormKind: grade = r!Paygrade: result = r!Title & " / " & kind & " / " & grade & vbCrLf: r.Close
    For Each key In ids.Keys
        Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & CLng(key), dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", "Selected member no longer exists."
        If r!IsArchived Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", "Choose active members."
        If ValidationWornPaygrade(Nz(r!RankRate, "")) <> grade Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", r!FullName & ": rank/rate must match group paygrade."
        If Nz(r!ReportName, "") = "" Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", r!FullName & ": enter the explicit official report name in the member profile first."
        If DirectoryHasPeriod(CLng(key), kind, startDate, endDate) Then Err.Raise vbObjectError + 6990, "modDirectoryBatch.PreviewDirectoryReports", r!FullName & ": a saved report already has this type and period."
        memberStamps.Add CStr(key), CLng(r!VersionCounter)
        result = result & r!FullName & " / " & r!RankRate & " / " & r!ReportName & vbCrLf: r.Close
    Next
    fromDate = DateValue(startDate): toDate = DateValue(endDate): pendingGroup = groupID
    groupBaseline = GroupStamp(groupID): defaultBaseline = UXDefaultsStamp(groupID): Set selectedMembers = ids
    result = result & vbCrLf & "Shared starting fields (blank means unentered):" & vbCrLf
    For Each field In UXSharedKeys()
        defaultValue = UXEffectiveDefault(groupID, CStr(field))
        If field = "PeriodStart" Then defaultValue = Format$(fromDate, "yyyy-mm-dd")
        If field = "PeriodEnd" Then defaultValue = Format$(toDate, "yyyy-mm-dd")
        result = result & UXFieldCaption(CStr(field), kind) & ": " & Nz(defaultValue, "(blank)") & vbCrLf
    Next
    PreviewDirectoryReports = result & vbCrLf & Format$(fromDate, "yyyy-mm-dd") & " to " & Format$(toDate, "yyyy-mm-dd") & vbCrLf & "Creates new drafts with global/group starting defaults. Existing reports stay unchanged."
End Function
Private Function DirectoryHasPeriod(ByVal memberID As Long, ByVal kind As String, ByVal startDate As Date, ByVal endDate As Date) As Boolean
    DirectoryHasPeriod = Nz(MajorScalar("SELECT Count(*) FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND r.FormKind=" & QuoteSQL(kind) & " AND r.PeriodStart=#" & Format$(startDate, "yyyy-mm-dd") & "# AND r.PeriodEnd=#" & Format$(endDate, "yyyy-mm-dd") & "#"), 0) > 0
End Function
Public Function ApplyDirectoryReports() As Long
    Dim w As DAO.Workspace, key As Variant, field As Variant, r As DAO.Recordset, values As Object, kind As String, grade As String, id As Long, failure As Long, message As String, op As Long
    If selectedMembers Is Nothing Then Err.Raise vbObjectError + 6991, "modDirectoryBatch.ApplyDirectoryReports", "Preview selected members first."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If GroupStamp(pendingGroup) <> groupBaseline Or UXDefaultsStamp(pendingGroup) <> defaultBaseline Then Err.Raise vbObjectError + 6991, "modDirectoryBatch.ApplyDirectoryReports", "Group or defaults changed. Preview again."
    kind = MajorScalar("SELECT FormKind FROM SummaryGroup WHERE GroupID=" & pendingGroup): grade = MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & pendingGroup)
    For Each key In selectedMembers.Keys
        Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & CLng(key), dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 6991, "modDirectoryBatch.ApplyDirectoryReports", "Member no longer exists. Preview again."
        If r!IsArchived Or r!VersionCounter <> memberStamps(CStr(key)) Then Err.Raise vbObjectError + 6991, "modDirectoryBatch.ApplyDirectoryReports", "Member profile changed. Preview again."
        If DirectoryHasPeriod(CLng(key), kind, fromDate, toDate) Then Err.Raise vbObjectError + 6991, "modDirectoryBatch.ApplyDirectoryReports", "A report for this type/period was saved since preview. Preview again."
        Set values = CreateObject("Scripting.Dictionary")
        For Each field In DraftFieldKeys(): values.Add CStr(field), Null: Next
        For Each field In UXSharedKeys(): values(CStr(field)) = UXEffectiveDefault(pendingGroup, CStr(field)): Next
        values("FormKind") = kind: values("PeriodStart") = fromDate: values("PeriodEnd") = toDate
        values("OfficialName") = r!ReportName: values("RankRate") = r!RankRate: values("MemberSSN") = r!MemberSSN
        DirectoryCreationContext CLng(key)
        id = CommitDraft(0, kind, kind & " - " & r!FullName, r!FullName, fromDate, toDate, "", "Create from command directory", values("OfficialComments"), _
            FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
            FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
        DirectoryCreationContext 0: r.Close
        AddGroupMember pendingGroup, id, grade: ApplyDirectoryReports = ApplyDirectoryReports + 1
    Next
    op = BeginOperation("Create reports from directory", "Group " & pendingGroup & "; member count " & selectedMembers.Count)
    FinishOperation op, "Completed", "New drafts and directory/group links committed atomically."
    w.CommitTrans: Set selectedMembers = Nothing: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modDirectoryBatch", "ApplyDirectoryReports", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    DirectoryCreationContext 0: SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
