Attribute VB_Name = "modTraits"
Option Compare Binary
Option Explicit
Private editorKind As String
Public Function TraitKeys() As Variant
    TraitKeys = Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39")
End Function
Public Function IsTraitKey(ByVal key As String) As Boolean
    Dim item As Variant
    For Each item In TraitKeys()
        If item = key Then IsTraitKey = True: Exit Function
    Next
End Function
Public Function ValidateTraits(ByVal values As Variant) As Variant
    Dim i As Long, mark As Variant, allowed As Boolean
    If Not IsArray(values) Then Err.Raise vbObjectError + 2000, , "Traits must contain seven marks."
    If LBound(values) <> 0 Or UBound(values) <> 6 Then Err.Raise vbObjectError + 2000, , "Traits must contain seven marks."
    For i = 0 To 6
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2001, , "Trait marks must be listed text values or blank, not numeric codes."
            allowed = False
            For Each mark In Array("", "NOB", "1.0", "2.0", "3.0", "4.0", "5.0")
                If values(i) = mark Then allowed = True
            Next
            If Not allowed Then Err.Raise vbObjectError + 2002, , "Block " & (33 + i) & ": choose NOB, 1.0, 2.0, 3.0, 4.0 or 5.0, or clear the mark."
        End If
    Next
    ValidateTraits = values
End Function
Public Function TraitsFromEditor() As Variant
    Dim values As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null)
    For i = 0 To 6: values(i) = Forms!frmEditor.Controls("cboTrait" & (33 + i)).Value: Next
    TraitsFromEditor = values
End Function
Public Function TraitsFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long, schema As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null): keys = TraitKeys()
    schema = RevisionSnapshotSchema(revisionID)
    For i = 0 To 6
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If schema >= 10 Then r.Close: Err.Raise vbObjectError + 2003, , "Trait snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2004, , "Trait snapshot has an inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    TraitsFromRevision = ValidateTraits(values)
End Function
Public Function InitializeTraitEditor() As Boolean
    Dim block As Long
    editorKind = Nz(Forms!frmEditor!cboKind, "EVAL")
    For block = 33 To 39
        Forms!frmEditor.Controls("lblTrait" & block).Caption = block & ". " & Replace(Replace(CatalogueTraitTitle(editorKind, block), vbCr, ""), vbLf, " ")
    Next
    Call ShowTraitGuidance(33)
    InitializeTraitEditor = True
End Function
Public Function EditorTypeChanged() As Boolean
    Dim values As Variant, value As Variant, entered As Boolean, block As Long
    values = TraitsFromEditor()
    For Each value In values
        If Len(Nz(value, "")) > 0 Then entered = True
    Next
    If editorKind <> Nz(Forms!frmEditor!cboKind, "") And entered Then
        If MsgBox("Trait meanings differ by report type. Change type and clear the seven pending marks? Cancel keeps the current type and marks.", vbOKCancel + vbQuestion, "Change draft type") <> vbOK Then
            Forms!frmEditor!cboKind = editorKind
        Else
            For block = 33 To 39: Forms!frmEditor.Controls("cboTrait" & block) = Null: Next
        End If
    End If
    Call InitializeTraitEditor(): Call RefreshNarrativeLabel()
    EditorTypeChanged = True
End Function
Public Function ShowTraitGuidance(ByVal block As Long) As Boolean
    Dim kind As String, mark As String, text As String
    kind = Nz(Forms!frmEditor!cboKind, "EVAL")
    mark = Nz(Forms!frmEditor.Controls("cboTrait" & block), "")
    text = block & ". " & CatalogueTraitTitle(kind, block) & vbCrLf & CatalogueTraitDescription(kind, block, "")
    If Len(mark) > 0 Then text = text & vbCrLf & vbCrLf & "Selected mark: " & mark & vbCrLf & CatalogueTraitDescription(kind, block, mark)
    text = text & vbCrLf & vbCrLf & "NAVFIT reference descriptions. Blank differs from NOB. Whole-report observation status and recommendation rules are not assessed."
    Forms!frmEditor!txtTraitGuidance = text: ShowTraitGuidance = True
    Call RefreshTraitSummary()
End Function
Public Function TraitSummary(ByVal values As Variant) As String
    Dim i As Long, numericCount As Long, nobCount As Long, blankCount As Long, total As Long, average As String
    On Error GoTo Invalid
    values = ValidateTraits(values)
    For i = 0 To 6
        If Len(Nz(values(i), "")) = 0 Then
            blankCount = blankCount + 1
        ElseIf values(i) = "NOB" Then
            nobCount = nobCount + 1
        Else
            total = total + CLng(Left$(values(i), 1)): numericCount = numericCount + 1
        End If
    Next
    If numericCount = 0 Then
        average = "Unavailable (no numeric marks)"
    Else
        average = Format$(CDbl(total) / numericCount, "0.00")
        If blankCount > 0 Then average = average & " (partial; blank marks remain)"
    End If
    TraitSummary = "Numeric average: " & average & "; numeric " & numericCount & "; NOB " & nobCount & "; blank " & blankCount
    Exit Function
Invalid:
    TraitSummary = "Numeric average unavailable: unsupported pending mark. Choose listed marks or clear the input."
End Function
Public Function RefreshTraitSummary() As Boolean
    Dim reportID As Long, revisionID As Long, revisionNumber As Long
    Forms!frmEditor!txtPendingTraitSummary = "Editor values: " & TraitSummary(TraitsFromEditor())
    reportID = Nz(Forms!frmEditor!txtReportID, 0)
    revisionID = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    If revisionID = 0 Then
        Forms!frmEditor!txtSavedTraitSummary = "Saved values: no committed revision."
    Else
        revisionNumber = Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & revisionID), 0)
        Forms!frmEditor!txtSavedTraitSummary = "Saved revision " & revisionNumber & ": " & TraitSummary(TraitsFromRevision(revisionID))
    End If
    RefreshTraitSummary = True
End Function
