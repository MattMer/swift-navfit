Attribute VB_Name = "modUXCalculationContext"
Option Compare Database
Option Explicit
Public Function UXTraitContext(ByVal values As Object) As String
    Dim key As Variant, total As Double, graded As Long, blank As Long, nob As Long
    If Nz(values("ObservationStatus"), "") = "NOB" Then
        UXTraitContext = "Whole-report NOB: all seven traits excluded from this average.": Exit Function
    End If
    For Each key In TraitKeys()
        Select Case Nz(values(CStr(key)), "")
            Case "1.0", "2.0", "3.0", "4.0", "5.0": graded = graded + 1: total = total + CDbl(values(CStr(key)))
            Case "NOB": nob = nob + 1
            Case Else: blank = blank + 1
        End Select
    Next
    UXTraitContext = "Trait sum " & Format$(total, "0.00") & " / " & graded & " graded traits"
    If graded > 0 Then UXTraitContext = UXTraitContext & " = " & Format$(total / graded, "0.00") Else UXTraitContext = UXTraitContext & ": unavailable"
    UXTraitContext = UXTraitContext & ". Excluded: " & nob & " NOB, " & blank & " blank/ungraded. Arithmetic only."
End Function
Public Function UXGroupCalculationText(ByVal groupID As Long, Optional ByVal loaded As DAO.Recordset = Nothing) As String
    Dim r As DAO.Recordset, values As Object, key As Variant, marks As Long, total As Double, archived As Long, nob As Long, ungraded As Long, reports As Long, baseline As DAO.Recordset
    If loaded Is Nothing Then Set r = RSGroupRows(groupID) Else Set r = loaded.Clone
    If Not (r.BOF And r.EOF) Then r.MoveFirst
    Do While Not r.EOF
        reports = reports + 1
        If r!IsArchived Then
            archived = archived + 1
        Else
            If Nz(r!ObservationStatus, "") = "NOB" Then
                nob = nob + 1
            Else
                For Each key In TraitKeys()
                    Select Case Nz(r.Fields(CStr(key)).Value, "")
                        Case "1.0", "2.0", "3.0", "4.0", "5.0": marks = marks + 1: total = total + CDbl(r.Fields(CStr(key)).Value)
                        Case Else: ungraded = ungraded + 1
                    End Select
                Next
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    UXGroupCalculationText = "Saved revisions only; pending report/comparison edits are excluded." & vbCrLf & "Group trait average: " & Format$(total, "0.00") & " trait sum / " & marks & " graded traits"
    If marks > 0 Then UXGroupCalculationText = UXGroupCalculationText & " = " & Format$(total / marks, "0.00") Else UXGroupCalculationText = UXGroupCalculationText & ": unavailable"
    UXGroupCalculationText = UXGroupCalculationText & vbCrLf & reports & " member reports; excluded " & archived & " archived reports, " & nob & " whole-report NOB reports, " & ungraded & " blank/NOB/ungraded traits. PSG RSCA requires a matching senior and paygrade baseline excluding this group."
    Set baseline = CurrentDb.OpenRecordset("SELECT b.* FROM RSCABaseline AS b INNER JOIN GroupRSCABaseline AS g ON b.BaselineID=g.BaselineID WHERE g.GroupID=" & groupID, dbOpenSnapshot)
    If Not baseline.EOF Then
        UXGroupCalculationText = UXGroupCalculationText & vbCrLf & "Manual historical baseline: " & baseline!Paygrade & ", as of " & Format$(baseline!AsOfDate, "yyyy-mm-dd") & ", " & baseline!TraitCount & " prior traits; " & IIf(baseline!TraitCount = 0, "zero prior traits", IIf(IsNull(baseline!TraitSum), "rounded RSCA estimate", "exact entered sum")) & ". Source: " & Nz(baseline!SourceNote, "No source note entered")
    End If
    baseline.Close

End Function
Public Function UXEditorCalculationDetails() As Boolean
    Dim text As String, id As Long, groupID As Long, values As Object, key As Variant
    Set values = CreateObject("Scripting.Dictionary")
    values.Add "ObservationStatus", UXSurfacePending("cboObservationStatus", Forms!frmEditor!cboObservationStatus.Value)
    For Each key In TraitKeys()
        values.Add CStr(key), UXSurfacePending(UXFieldControl(CStr(key)), Forms!frmEditor.Controls(UXFieldControl(CStr(key))).Value)
    Next
    id = Nz(Forms!frmEditor!txtReportID, 0)
    text = "Current editor values (pending): " & UXTraitContext(values) & vbCrLf & vbCrLf & "Latest saved revision: " & UXTraitContext(SavedFields(id))
    groupID = Nz(Forms!frmEditor!uxGroup, 0)
    If groupID > 0 Then text = text & vbCrLf & vbCrLf & UXGroupCalculationText(groupID)
    MsgBox text, vbInformation, "Calculation details"
    UXEditorCalculationDetails = True
End Function
