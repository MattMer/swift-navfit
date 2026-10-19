Attribute VB_Name = "GroupReadLegacy"
Option Compare Binary
Option Explicit
Public Function LegacyUXGroupAverage(ByVal groupID As Long) As Variant
    Dim r As DAO.Recordset, key As Variant, value As Variant, total As Double, marks As Long
    Set r = CurrentDb.OpenRecordset("SELECT r.CurrentRevisionID,r.ObservationStatus FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        If Nz(r!ObservationStatus, "") <> "NOB" Then
            For Each key In TraitKeys()
                value = MajorRevisionValue(Nz(r!CurrentRevisionID, 0), CStr(key))
                Select Case Nz(value, "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(value): marks = marks + 1
                End Select
            Next
        End If
        r.MoveNext
    Loop
    r.Close
    If marks > 0 Then LegacyUXGroupAverage = total / marks Else LegacyUXGroupAverage = Null
End Function
Public Function LegacyRSGroupPSG(ByVal groupID As Long) As Object
    Dim result As Object, r As DAO.Recordset, members As DAO.Recordset, senior As Object, values As Object, key As Variant, count As Long, total As Double, previousCount As Long, previousSum As Double, memberCount As Long
    Set result = CreateObject("Scripting.Dictionary"): result.Add "Value", Null: result.Add "Reason", "Choose a reporting senior and historical baseline.": result.Add "Estimated", False: result.Add "Partial", False
    Set r = CurrentDb.OpenRecordset("SELECT b.*,g.ExcludesGroup,g.SeniorID AS SelectedSeniorID FROM GroupRSCABaseline AS g INNER JOIN RSCABaseline AS b ON g.BaselineID=b.BaselineID WHERE g.GroupID=" & groupID, dbOpenSnapshot)
    If r.EOF Then r.Close: Set LegacyRSGroupPSG = result: Exit Function
    If Not r!ExcludesGroup Then result("Reason") = "Confirm that historical totals exclude the current group.": r.Close: Set LegacyRSGroupPSG = result: Exit Function
    If r!Paygrade <> MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupID) Then result("Reason") = "Historical baseline paygrade does not match this group.": r.Close: Set LegacyRSGroupPSG = result: Exit Function
    If r!SelectedSeniorID <> r!SeniorID Then result("Reason") = "Baseline belongs to a different reporting senior.": r.Close: Set LegacyRSGroupPSG = result: Exit Function
    Set senior = RSSeniorValues(r!SeniorID)
    previousCount = r!TraitCount
    RSValidateBaseline CStr(r!Paygrade), r!CurrentRSCA.Value, previousCount, r!TraitSum.Value
    If previousCount = 0 Then
        previousSum = 0
    ElseIf IsNull(r!TraitSum) Then
        previousSum = CDbl(r!CurrentRSCA) * previousCount: result("Estimated") = True
    Else
        previousSum = r!TraitSum
    End If
    r.Close
    Set members = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID, dbOpenSnapshot)
    Do While Not members.EOF
        Set values = SavedFields(members!ReportID)
        If Len(Nz(senior("ReportingSeniorSSN"), "")) = 0 Or Nz(values("ReportingSeniorSSN"), "") <> Nz(senior("ReportingSeniorSSN"), "") Then
            result("Reason") = "All group reports must use the selected reporting senior SSN. Apply senior details first.": members.Close: Set LegacyRSGroupPSG = result: Exit Function
        End If
        If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & members!ReportID), True) Then
            result("Reason") = "Remove or recover archived group members before calculating PSG RSCA.": members.Close: Set LegacyRSGroupPSG = result: Exit Function
        End If
        If Nz(values("ObservationStatus"), "") <> "NOB" Then
            memberCount = 0
            For Each key In TraitKeys()
                Select Case Nz(values(CStr(key)), "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(values(CStr(key))): count = count + 1: memberCount = memberCount + 1
                End Select
            Next
            If memberCount < 7 Then result("Partial") = True
        End If
        members.MoveNext
    Loop
    members.Close
    If count = 0 Then result("Reason") = "No graded traits are available.": Set LegacyRSGroupPSG = result: Exit Function
    result("Value") = (previousSum + total) / (CDbl(previousCount) + count)
    result("Reason") = "(" & Format$(previousSum, "0.00") & " historical sum + " & total & " group sum) / (" & previousCount & " historical traits + " & count & " group traits)"
    Set LegacyRSGroupPSG = result
End Function
Public Function LegacyUXGroupCalculationText(ByVal groupID As Long) As String
    Dim r As DAO.Recordset, values As Object, key As Variant, marks As Long, total As Double, archived As Long, nob As Long, ungraded As Long, reports As Long, baseline As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT r.* FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID, dbOpenSnapshot)
    Do While Not r.EOF
        reports = reports + 1
        If r!IsArchived Then
            archived = archived + 1
        Else
            Set values = SavedFields(r!ReportID)
            If Nz(values("ObservationStatus"), "") = "NOB" Then
                nob = nob + 1
            Else
                For Each key In TraitKeys()
                    Select Case Nz(values(CStr(key)), "")
                        Case "1.0", "2.0", "3.0", "4.0", "5.0": marks = marks + 1: total = total + CDbl(values(CStr(key)))
                        Case Else: ungraded = ungraded + 1
                    End Select
                Next
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    LegacyUXGroupCalculationText = "Saved revisions only; pending report/comparison edits are excluded." & vbCrLf & "Group trait average: " & Format$(total, "0.00") & " trait sum / " & marks & " graded traits"
    If marks > 0 Then LegacyUXGroupCalculationText = LegacyUXGroupCalculationText & " = " & Format$(total / marks, "0.00") Else LegacyUXGroupCalculationText = LegacyUXGroupCalculationText & ": unavailable"
    LegacyUXGroupCalculationText = LegacyUXGroupCalculationText & vbCrLf & reports & " member reports; excluded " & archived & " archived reports, " & nob & " whole-report NOB reports, " & ungraded & " blank/NOB/ungraded traits. PSG RSCA requires a matching senior and paygrade baseline excluding this group."
    Set baseline = CurrentDb.OpenRecordset("SELECT b.* FROM RSCABaseline AS b INNER JOIN GroupRSCABaseline AS g ON b.BaselineID=g.BaselineID WHERE g.GroupID=" & groupID, dbOpenSnapshot)
    If Not baseline.EOF Then
        LegacyUXGroupCalculationText = LegacyUXGroupCalculationText & vbCrLf & "Manual historical baseline: " & baseline!Paygrade & ", as of " & Format$(baseline!AsOfDate, "yyyy-mm-dd") & ", " & baseline!TraitCount & " prior traits; " & IIf(baseline!TraitCount = 0, "zero prior traits", IIf(IsNull(baseline!TraitSum), "rounded RSCA estimate", "exact entered sum")) & ". Source: " & Nz(baseline!SourceNote, "No source note entered")
    End If
    baseline.Close

End Function
