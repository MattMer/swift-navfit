Attribute VB_Name = "modTraits"
Option Compare Binary
Option Explicit
Private editorKind As String
'The reviewed JSON seeds fresh databases; all runtime wording comes from this
'versioned reference table. Unknown callers retain the prior empty-result contract.
Public Function TraitDefinitionValue(ByVal kind As String, ByVal block As Long, ByVal mark As String, ByVal title As Boolean) As String
    Dim r As DAO.Recordset
    If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Exit Function
    If block < 33 Or block > 39 Then Exit Function
    If mark <> "" And mark <> "NOB" And mark <> "1.0" And mark <> "2.0" And mark <> "3.0" And mark <> "4.0" And mark <> "5.0" Then Exit Function
    If mark = "" Then mark = "GENERAL"
    Set r = CurrentDb.OpenRecordset("SELECT TraitTitle,DescriptionText,SourceSHA256 FROM TraitDefinition WHERE FormKind='" & kind & "' AND BlockNumber=" & block & " AND Mark='" & mark & "' AND CatalogueVersion=" & TRAIT_CATALOGUE_VERSION, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 9300, "modTraits.TraitDefinitionValue", "Required trait definition is missing. Restore a verified application copy."
    If StrComp(Nz(r!SourceSHA256, ""), TraitCatalogueHash(), vbBinaryCompare) <> 0 Then r.Close: Err.Raise vbObjectError + 9301, "modTraits.TraitDefinitionValue", "Trait catalogue source identity differs from this application."
    If title Then TraitDefinitionValue = Nz(r!TraitTitle, "") Else TraitDefinitionValue = Nz(r!DescriptionText, "")
    r.Close
End Function

'Called inside the production import transaction. Retain every older version and
'reject conflicting identities/wording rather than overwriting either catalogue.
Public Sub ImportTraitDefinitions(ByVal source As DAO.Database, ByVal target As DAO.Database)
    Dim rows As DAO.Recordset, existing As DAO.Recordset, field As DAO.field, sql As String
    If source.TableDefs("TraitDefinition").Fields.Count <> target.TableDefs("TraitDefinition").Fields.Count Then Err.Raise vbObjectError + 9303, "modTraits.ImportTraitDefinitions", "Invalid trait reference column inventory. No workspace data imported."
    For Each field In target.TableDefs("TraitDefinition").Fields
        If source.TableDefs("TraitDefinition").Fields(field.Name).Type <> field.Type Then Err.Raise vbObjectError + 9303, "modTraits.ImportTraitDefinitions", "Invalid trait reference column types. No workspace data imported."
    Next
    Set rows = source.OpenRecordset("TraitDefinition", dbOpenSnapshot)
    Do While Not rows.EOF
        sql = "SELECT * FROM TraitDefinition WHERE FormKind=" & QuoteSQL(rows!formKind) & " AND BlockNumber=" & CLng(rows!BlockNumber) & " AND Mark=" & QuoteSQL(rows!mark) & " AND CatalogueVersion=" & CLng(rows!CatalogueVersion)
        Set existing = target.OpenRecordset(sql, dbOpenDynaset)
        If existing.EOF Then
            existing.AddNew
            For Each field In rows.fields: existing.fields(field.name).value = field.value: Next
            existing.Update
            existing.Bookmark = existing.LastModified
        End If
        For Each field In rows.Fields
            If Not SameFieldValue(existing.Fields(field.Name).Value, field.Value) Then Err.Raise vbObjectError + 9302, "modTraits.ImportTraitDefinitions", "Conflicting trait reference data. No workspace data imported."
        Next
        existing.Close: rows.MoveNext
    Loop
    rows.Close
End Sub
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
    If Not IsArray(values) Then Err.Raise vbObjectError + 2000, "modTraits.ValidateTraits", "Traits must contain seven marks."
    If LBound(values) <> 0 Or UBound(values) <> 6 Then Err.Raise vbObjectError + 2000, "modTraits.ValidateTraits", "Traits must contain seven marks."
    For i = 0 To 6
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2001, "modTraits.ValidateTraits", "Trait marks must be listed text values or blank, not numeric codes."
            allowed = False
            For Each mark In Array("", "NOB", "1.0", "2.0", "3.0", "4.0", "5.0")
                If values(i) = mark Then allowed = True
            Next
            If Not allowed Then Err.Raise vbObjectError + 2002, "modTraits.ValidateTraits", "Block " & (33 + i) & ": choose NOB, 1.0, 2.0, 3.0, 4.0 or 5.0, or clear the mark."
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
            If schema >= 10 Then r.Close: Err.Raise vbObjectError + 2003, "modTraits.TraitsFromRevision", "Trait snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2004, "modTraits.TraitsFromRevision", "Trait snapshot has an inconsistent null flag."
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
