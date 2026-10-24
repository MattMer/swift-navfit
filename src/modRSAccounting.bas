Attribute VB_Name = "modRSAccounting"
Option Explicit

'Accounting is append-only. Active report claims prevent counting one report twice;
'reopening/archiving does not release claims or subtract historical totals.
Public Function RSPosting(ByVal groupID As Long) As Long
    RSPosting = Nz(MajorScalar("SELECT CompletionID FROM GroupCompletionCurrent WHERE GroupID=" & groupID), 0)
End Function

Public Function RSPostingRevisionsMatch(ByVal groupID As Long) As Boolean
    Dim id As Long, r As DAO.Recordset
    id = RSPosting(groupID): If id = 0 Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT c.ReportID,c.RevisionID,r.CurrentRevisionID FROM GroupCompletionReport AS c INNER JOIN Report AS r ON c.ReportID=r.ReportID WHERE c.CompletionID=" & id, dbOpenSnapshot)
    RSPostingRevisionsMatch = True
    Do While Not r.EOF
        If Nz(r!RevisionID, 0) <> Nz(r!CurrentRevisionID, 0) Then RSPostingRevisionsMatch = False: Exit Do
        r.MoveNext
    Loop
    r.Close
End Function

Public Function RSReconcileAccounting() As String
    Dim r As DAO.Recordset, db As DAO.Database, rows As Long, legacy As Long
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT p.CompletionID,p.TraitSum,p.TraitCount,Sum(c.TraitSum) AS ChildSum,Sum(c.TraitCount) AS ChildCount FROM GroupCompletion AS p LEFT JOIN GroupCompletionReport AS c ON p.CompletionID=c.CompletionID WHERE p.EvidenceKind='Saved revisions' GROUP BY p.CompletionID,p.TraitSum,p.TraitCount", dbOpenSnapshot)
    Do While Not r.EOF
        If CDbl(r!TraitSum) <> CDbl(Nz(r!ChildSum, 0)) Or CLng(r!TraitCount) <> CLng(Nz(r!ChildCount, 0)) Then Err.Raise vbObjectError + 4865, "modRSAccounting.RSReconcileAccounting", "Posting child totals do not reconcile: " & r!CompletionID
        rows = rows + 1: r.MoveNext
    Loop
    r.Close
    legacy = PersistenceCount("GroupCompletion", "EvidenceKind<>'Saved revisions'")
    RSReconcileAccounting = "PASS: " & rows & " revision-backed postings reconcile; " & legacy & " legacy postings retain aggregate totals with unknown revision evidence."
End Function

Public Function RSPostingValue(ByVal groupID As Long, ByVal key As String) As String
    Dim id As Long, r As DAO.Recordset
    id = RSPosting(groupID): If id = 0 Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM GroupCompletion WHERE CompletionID=" & id, dbOpenSnapshot)
    Select Case key
        Case "RSCAPostedBaseline": RSPostingValue = CStr(r!PostedBaselineID)
        Case "RSCAGroupSum": RSPostingValue = CStr(r!TraitSum)
        Case "RSCAGroupCount": RSPostingValue = CStr(r!TraitCount)
        Case "RSCACompletedAt": RSPostingValue = Format$(r!CompletedAt, "yyyy-mm-dd hh:nn:ss")
    End Select
    r.Close
    If key = "RSCAReportIDs" Then
        RSPostingValue = "|"
        Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupCompletionReport WHERE CompletionID=" & id & " ORDER BY ReportID", dbOpenSnapshot)
        Do While Not r.EOF: RSPostingValue = RSPostingValue & r!ReportID & "|": r.MoveNext: Loop
        r.Close
    End If
End Function

Public Sub RSWritePosting(ByVal groupID As Long, ByVal seniorID As Long, ByVal parentBaseline As Long, ByVal baseline As Long, ByVal total As Double, ByVal count As Long, ByVal rows As DAO.Recordset)
    Dim db As DAO.Database, p As DAO.Recordset, child As DAO.Recordset, id As Long, old As Long
    Dim key As Variant, marks As Long, sum As Double, value As String
    Set db = CurrentDb: old = RSPosting(groupID)
    Set p = db.OpenRecordset("GroupCompletion", dbOpenDynaset)
    p.AddNew: p!GroupID = groupID: p!SeniorID = seniorID: p!ParentBaselineID = parentBaseline
    p!PostedBaselineID = baseline: p!TraitSum = total: p!TraitCount = count: p!CompletedAt = Now()
    p!EvidenceKind = "Saved revisions": If old > 0 Then p!PriorCompletionID = old
    p.Update: p.Bookmark = p.LastModified: id = p!CompletionID: p.Close
    db.Execute "DELETE FROM RSCountedReport WHERE CompletionID=" & old, dbFailOnError
    Set child = db.OpenRecordset("GroupCompletionReport", dbOpenDynaset)
    If Not rows.BOF Or Not rows.EOF Then rows.MoveFirst
    Do While Not rows.EOF
        If Nz(rows!ObservationStatus, "") <> "NOB" Then
            marks = 0: sum = 0
            For Each key In TraitKeys()
                value = Nz(rows.Fields(CStr(key)).Value, "")
                If value <> "NOB" Then marks = marks + 1: sum = sum + CDbl(value)
            Next
            child.AddNew: child!CompletionID = id: child!ReportID = rows!ReportID
            child!RevisionID = rows!CurrentRevisionID: child!RevisionOwnerReportID = rows!ReportID
            child!TraitCount = marks: child!TraitSum = sum: child.Update
            db.Execute "INSERT INTO RSCountedReport (ReportID,CompletionID) VALUES (" & rows!ReportID & "," & id & ")", dbFailOnError
        End If
        rows.MoveNext
    Loop
    child.Close
    db.Execute "DELETE FROM GroupCompletionCurrent WHERE GroupID=" & groupID, dbFailOnError
    db.Execute "INSERT INTO GroupCompletionCurrent (GroupID,CompletionID) VALUES (" & groupID & "," & id & ")", dbFailOnError
End Sub

Public Sub RSGuardPostedSenior(ByVal groupID As Long, ByVal seniorID As Long)
    Dim id As Long
    id = RSPosting(groupID)
    If id > 0 Then
        If CLng(DLookup("SeniorID", "GroupCompletion", "CompletionID=" & id)) <> seniorID Then Err.Raise vbObjectError + 4862, "modRSAccounting.RSGuardPostedSenior", "A counted group cannot be reassigned to another reporting senior. Reconcile its accounting history first."
    End If
End Sub

'Called only inside the production import transaction, after original-field verification.
'Old XML remains untouched as source evidence. A legacy report's posting revision
'was never recorded: leave it Null rather than attribute it to today's revision.
Public Sub MigrateRSAccounting()
    Dim db As DAO.Database, r As DAO.Recordset, p As DAO.Recordset, c As DAO.Recordset
    Dim d As Object, node As Object, id As Variant, groupID As Long, baseline As Long, parent As Long
    Dim posting As Long, note As String, state As String, senior As Long
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT * FROM RSCABaseline ORDER BY BaselineID", dbOpenDynaset)
    Do While Not r.EOF
        note = Nz(r!SourceNote, "")
        If Left$(note, 25) = "Swift local RSCA; parent=" Then
            parent = Val(Mid$(note, 26))
            If parent <= 0 Or parent >= r!BaselineID Then Err.Raise vbObjectError + 4863, "modRSAccounting.MigrateRSAccounting", "Invalid legacy baseline ancestry. Import rolled back."
            If PersistenceCount("RSCABaseline", "BaselineID=" & parent & " AND SeniorID=" & r!SeniorID & " AND Paygrade=" & QuoteSQL(CStr(r!Paygrade))) <> 1 Then Err.Raise vbObjectError + 4863, "modRSAccounting.MigrateRSAccounting", "Legacy baseline ancestor does not match. Import rolled back."
            If IsNull(r!ParentBaselineID) Then r.Edit: r!ParentBaselineID = parent: r.Update
        End If
        r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT GroupID,KeySpec FROM SummaryGroup", dbOpenSnapshot)
    Do While Not r.EOF
        groupID = r!GroupID
        Set d = CreateObject("MSXML2.DOMDocument.6.0"): d.async = False: d.resolveExternals = False: d.setProperty "ProhibitDTD", True
        If Len(Nz(r!KeySpec, "")) > 0 Then
            If Not d.LoadXML(CStr(r!KeySpec)) Then Err.Raise vbObjectError + 4864, "modRSAccounting.MigrateRSAccounting", "Invalid legacy group policy. Import rolled back."
            Set node = d.selectSingleNode("/SwiftGroupPolicy/Workflow")
            If Not node Is Nothing And PersistenceCount("GroupLifecycle", "GroupID=" & groupID) = 0 Then
                Set p = db.OpenRecordset("GroupLifecycle", dbOpenDynaset)
                p.AddNew: p!GroupID = groupID: p!WorkflowState = node.Text: p!LegacyState = node.Text
                Set node = d.selectSingleNode("/SwiftGroupPolicy/WorkflowBy")
                If Not node Is Nothing Then p!ChangedBy = CLng(node.Text)
                Set node = d.selectSingleNode("/SwiftGroupPolicy/WorkflowAt")
                If Not node Is Nothing Then p!ChangedAt = CDate(node.Text)
                p.Update: p.Close
            End If
            Set node = d.selectSingleNode("/SwiftGroupPolicy/RSCAPostedBaseline")
            If Not node Is Nothing And RSPosting(groupID) = 0 Then
                baseline = CLng(node.Text)
                Set p = db.OpenRecordset("SELECT * FROM RSCABaseline WHERE BaselineID=" & baseline, dbOpenSnapshot)
                If p.EOF Then Err.Raise vbObjectError + 4864, "modRSAccounting.MigrateRSAccounting", "Missing legacy posted baseline. Import rolled back."
                senior = p!SeniorID: parent = Nz(p!ParentBaselineID, 0): p.Close
                If parent = 0 Then Err.Raise vbObjectError + 4864, "modRSAccounting.MigrateRSAccounting", "Missing legacy posting lineage. Import rolled back."
                Set p = db.OpenRecordset("GroupCompletion", dbOpenDynaset)
                p.AddNew: p!GroupID = groupID: p!SeniorID = senior: p!ParentBaselineID = parent: p!PostedBaselineID = baseline
                p!TraitSum = CDbl(d.selectSingleNode("/SwiftGroupPolicy/RSCAGroupSum").Text)
                p!TraitCount = CLng(d.selectSingleNode("/SwiftGroupPolicy/RSCAGroupCount").Text)
                p!CompletedAt = CDate(d.selectSingleNode("/SwiftGroupPolicy/RSCACompletedAt").Text)
                p!EvidenceKind = "Legacy totals; revision unknown"
                p.Update: p.Bookmark = p.LastModified: posting = p!CompletionID: p.Close
                Set c = db.OpenRecordset("GroupCompletionReport", dbOpenDynaset)
                For Each id In Split(d.selectSingleNode("/SwiftGroupPolicy/RSCAReportIDs").Text, "|")
                    If Len(id) > 0 Then
                        c.AddNew: c!CompletionID = posting: c!ReportID = CLng(id): c.Update
                        db.Execute "INSERT INTO RSCountedReport (ReportID,CompletionID) VALUES (" & CLng(id) & "," & posting & ")", dbFailOnError
                    End If
                Next
                c.Close
                db.Execute "INSERT INTO GroupCompletionCurrent (GroupID,CompletionID) VALUES (" & groupID & "," & posting & ")", dbFailOnError
            End If
        End If
        r.MoveNext
    Loop
    r.Close
End Sub
