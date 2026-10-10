Attribute VB_Name = "UXDemoSeed"
Option Compare Binary
Option Explicit
Private Function DemoReport(ByVal kind As String, ByVal member As String, ByVal rank As String, ByVal number As Long, ByVal senior As Object) As Long
    Dim id As Long, changes As Object, key As Variant, marks As Variant, i As Long
    id = CommitDraft(0, kind, "DEMO - " & member, member, #1/1/2026#, #9/30/2026#, "Synthetic practice report. Change any values to explore the app.", "Create synthetic demo", "SYNTHETIC EXAMPLE ONLY. Led a practice team, completed training milestones and supported a fictional command. Replace this text while exploring the editor.")
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "CommandEmployment", "SYNTHETIC COMMAND: Completed training exercises and fictional deployment milestones."
    changes.Add "OfficialName", UCase$(member): changes.Add "RankRate", rank
    changes.Add "MemberSSN", "0000000" & Format$(number, "00")
    changes.Add "UnitUIC", "12345": changes.Add "ShipStation", "DEMO COMMAND"
    changes.Add "DateReported", "2026-01-01": changes.Add "PromotionStatus", "REGULAR"
    changes.Add "ComponentActive", "Yes": changes.Add "ComponentTAR", "No": changes.Add "ComponentInactive", "No": changes.Add "ComponentATADSW", "No"
    changes.Add "OccasionPeriodic", "Yes": changes.Add "OccasionDetachment", "No": changes.Add "OccasionFrocking", "No": changes.Add "OccasionSpecial", "No"
    changes.Add "ReportRegular", "Yes": changes.Add "ReportConcurrent", "No": changes.Add "ReportOperationalCommander", "No"
    changes.Add "ObservationStatus", "Observed": changes.Add "PrimaryDuty", "TEAM LEADER"
    changes.Add "DutyDescription", "Synthetic duties: coordinate a fictional team, mentor members and track practice qualifications."
    changes.Add "CounselingDate", "2026-06-01": changes.Add "CounselorName", "DEMO COUNSELOR"
    changes.Add "PromotionRecommendation", "Promotable"
    marks = Array("4.0", "3.0", "4.0", "4.0", "3.0", "4.0", "4.0")
    For i = 0 To 6
        If number Mod 2 = 0 Then marks(i) = "4.0"
        changes.Add "Trait" & (33 + i), marks(i)
    Next
    For Each key In senior.Keys: changes.Add CStr(key), senior(key): Next
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Populate synthetic blocks")
    DemoReport = id
End Function
Public Function SeedUXDemo() As String
    Dim table As DAO.TableDef, db As DAO.Database, senior As Object, seniorID As Long, key As Variant
    Dim ids(1 To 9) As Long, folders(1 To 3) As Long, groups(1 To 3) As Long, baselines(1 To 3) As Long
    Dim author As Long, reviewer As Long, route As Long, template As Long, task As Long, changes As Object, i As Long
    Set db = CurrentDb
    'Never add demonstration data to a populated or operational workspace.
    For Each table In db.TableDefs
        If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" Then
            If DCount("*", "[" & table.Name & "]") > 0 Then Err.Raise vbObjectError + 5100, , "Demo seeding requires a completely blank workspace."
        End If
    Next
    ConfigureWorkspace "SWIFT NAVFIT DEMO - SYNTHETIC DATA", "Alex Morgan (demo)"
    author = MajorAuthor(): reviewer = AddContributor("Robin Chen (demo)")
    Set senior = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys(): senior.Add CStr(key), Null: Next
    senior("ReportingSeniorName") = "DEMO A": senior("ReportingSeniorLastName") = "DEMO": senior("ReportingSeniorFirstInitial") = "A"
    senior("ReportingSeniorGrade") = "CDR": senior("ReportingSeniorDesignator") = "1110": senior("ReportingSeniorTitle") = "CO"
    senior("ReportingSeniorUIC") = "12345": senior("ReportingSeniorSSN") = "000000001"
    senior("ReportingSeniorAddress1") = "100 DEMO WAY": senior("ReportingSeniorCity") = "EXAMPLE": senior("ReportingSeniorState") = "VA"
    senior("ReportingSeniorZip") = "00000": senior("ReportingSeniorPhone") = "2025550100": senior("ReportingSeniorDSN") = "5550100"
    seniorID = RSSaveSenior(0, "CDR A. Demo - fictional command", senior)
    baselines(1) = RSAddBaseline(seniorID, "E5", #12/31/2025#, 4.2, 100, 420, "Synthetic historical totals; current group excluded")
    baselines(2) = RSAddBaseline(seniorID, "E7", #12/31/2025#, 4.3, 100, 430, "Synthetic historical totals; current group excluded")
    baselines(3) = RSAddBaseline(seniorID, "O3", #12/31/2025#, 4.1, 100, 410, "Synthetic historical totals; current group excluded")
    ids(1) = DemoReport("EVAL", "Jamie Taylor", "ET2", 1, senior)
    ids(2) = DemoReport("EVAL", "Casey Morgan", "IT2", 2, senior)
    ids(3) = DemoReport("EVAL", "Sam Rivera", "LS2", 3, senior)
    ids(4) = DemoReport("CHIEFEVAL", "Robin Chen", "ETC", 4, senior)
    ids(5) = DemoReport("CHIEFEVAL", "Jordan Reed", "ITC", 5, senior)
    ids(6) = DemoReport("FITREP", "Avery Patel", "LT", 6, senior)
    ids(7) = DemoReport("FITREP", "Riley Brooks", "LT", 7, senior)
    ids(8) = DemoReport("EVAL", "Quinn Park", "YN2", 8, senior)
    ids(9) = CommitDraft(0, "EVAL", "DEMO - Incomplete draft", "Drew Blake", Null, Null, "Try entering a name, SSN and report period, then save.", "Synthetic incomplete draft")
    folders(1) = CreateFolder("E1-E6 demo reports"): folders(2) = CreateFolder("Chief demo reports"): folders(3) = CreateFolder("Officer demo reports")
    For i = 1 To 3: AssignFolder ids(i), folders(1): Next
    For i = 4 To 5: AssignFolder ids(i), folders(2): Next
    For i = 6 To 7: AssignFolder ids(i), folders(3): Next
    AssignFolder ids(8), folders(1)
    groups(1) = CreateSummaryGroup("DEMO E5 - September 2026", "EVAL", "E5")
    groups(2) = CreateSummaryGroup("DEMO E7 - September 2026", "CHIEFEVAL", "E7")
    groups(3) = CreateSummaryGroup("DEMO O3 - September 2026", "FITREP", "O3")
    For i = 1 To 3: AddGroupMember groups(1), ids(i), "E5", "Synthetic group": Next
    For i = 4 To 5: AddGroupMember groups(2), ids(i), "E7", "Synthetic group": Next
    For i = 6 To 7: AddGroupMember groups(3), ids(i), "O3", "Synthetic group": Next
    For i = 1 To 3
        RSSetGroupBaseline groups(i), seniorID, baselines(i), True
        SetFieldDefault "Group", groups(i), "ShipStation", "DEMO COMMAND"
    Next
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "ObservationStatus", "NOB"
    For Each key In TraitKeys(): changes.Add CStr(key), "NOB": Next
    Call CommitSavedFields(ids(3), CurrentRevision(ids(3)), changes, "Synthetic whole-report NOB example")
    task = AssignReview(ids(2), CurrentRevision(ids(2)), author, #10/15/2026#, "Synthetic task: check counseling and trait marks.")
    task = AssignReview(ids(6), CurrentRevision(ids(6)), author, #10/1/2026#, "Synthetic stale review: open the report and restart review for its current revision.")
    changes.RemoveAll: changes.Add "DraftNotes", "Changed after assignment so this review is intentionally stale."
    Call CommitSavedFields(ids(6), CurrentRevision(ids(6)), changes, "Synthetic stale-review example")
    task = AssignReview(ids(1), CurrentRevision(ids(1)), reviewer, #10/20/2026#, "Synthetic task assigned to the second contributor.")
    template = CreateRouteTemplate("DEMO local review", CStr(author), "Check the saved synthetic report.")
    route = StartRoute(ids(5), template, groups(2))
    DecideRouteStep route, "Accepted", "Synthetic route complete"
    ApproveSavedRoute route, "Synthetic local approval - reopen to edit"
    ArchiveDraft ids(8), True
    Call SaveSnippet(0, "DEMO - Mentoring", "SYNTHETIC EXAMPLE: Mentored a fictional team and completed practice training milestones.", "Synthetic UX demonstration", "ALL")
    Call AddReviewComment(ids(2), CurrentRevision(ids(2)), "Synthetic feedback: confirm the counseling date before finishing review.")
    template = CreateRouteTemplate("DEMO ordered review", author & "," & reviewer, "Review the saved blocks, then hand off to the next contributor.")
    route = StartRoute(ids(4), template)
    route = StartRoute(ids(7), template)
    Call AddReviewComment(ids(7), CurrentRevision(ids(7)), "Earlier-revision feedback: check trait marks before restarting review.")
    changes.RemoveAll: changes.Add "DraftNotes", "Changed after route creation; requester can restart this ordered review."
    Call CommitSavedFields(ids(7), CurrentRevision(ids(7)), changes, "Synthetic stale ordered-route example")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(ids(1))
    Forms!frmEditor!txtNotes = "RECOVERY DEMO: These pending notes have not been committed. Recover or discard this checkpoint to explore the workflow."
    If Not CheckpointEditor() Then Err.Raise vbObjectError + 5101, , "Demo recovery checkpoint could not be created."
    ' Intentional crash fixture: retain pending checkpoint without a normal exit decision.
    Forms!frmEditor.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
    db.Execute "UPDATE AppMetadata SET CompatibilityStatus='Synthetic UX demo - NAVFIT exchange remains unverified'", dbFailOnError
    SeedUXDemo = "PASS: 9 synthetic reports, 3 groups, 2 contributors, 1 reporting senior, 3 RSCA baselines, review/approval/archive/NOB/recovery examples."
End Function
