Attribute VB_Name = "modNAVFITAdapter"
Option Compare Binary
Option Explicit
Public Function NAVFITKind(ByVal value As Variant) As String
    Select Case Nz(value, "")
        Case "Eval": NAVFITKind = "EVAL"
        Case "Chief": NAVFITKind = "CHIEFEVAL"
        Case "FitRep": NAVFITKind = "FITREP"
        Case Else: Err.Raise vbObjectError + 5100, , "Unsupported or null NAVFIT report type."
    End Select
End Function
Public Function NAVFITType(ByVal kind As String) As String
    Select Case kind
        Case "EVAL": NAVFITType = "Eval"
        Case "CHIEFEVAL": NAVFITType = "Chief"
        Case "FITREP": NAVFITType = "FitRep"
        Case Else: Err.Raise vbObjectError + 5100, , "Unsupported Swift report type."
    End Select
End Function
Public Function NAVFITDecode(ByVal value As Variant, ByVal conversion As String, Optional ByVal century As Long = 0) As Variant
    Dim n As Long, s As String, y As Long, m As Long, dayValue As Long, months As Variant, i As Long, d As Date, choices As Variant
    If IsNull(value) Then NAVFITDecode = Null: Exit Function
    Select Case conversion
        Case "text": NAVFITDecode = CStr(value)
        Case "date"
            If VarType(value) <> vbDate Then Err.Raise vbObjectError + 5101, , "Native date is not a typed date."
            If TimeValue(value) <> TimeSerial(0, 0, 0) Then Err.Raise vbObjectError + 5101, , "Native date contains a time that Swift's date-only field cannot preserve. Review the source value."
            NAVFITDecode = Format$(value, "yyyy-mm-dd")
        Case "flag", "observation"
            If VarType(value) <> vbBoolean Then Err.Raise vbObjectError + 5102, , "Native checkbox is not Boolean."
            If conversion = "flag" Then
                If value Then NAVFITDecode = "Yes" Else NAVFITDecode = "No"
            Else
                If value Then NAVFITDecode = "NOB" Else NAVFITDecode = "Observed"
            End If
        Case "trait", "promotion"
            If Not IsNumeric(value) Then Err.Raise vbObjectError + 5103, , "Unknown native assessment code."
            If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < 0 Or CDbl(value) > 5 Then Err.Raise vbObjectError + 5103, , "Unknown native assessment code."
            n = CLng(value)
            If conversion = "trait" Then
                If n = 0 Then NAVFITDecode = "NOB" Else NAVFITDecode = CStr(n) & ".0"
            Else
                choices = Array("NOB", "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote"): NAVFITDecode = choices(n)
            End If
        Case "counseling"
            s = CStr(value)
            If s = "" Then NAVFITDecode = "": Exit Function
            If IsCounselingStatus(s) Then NAVFITDecode = s: Exit Function
            If IsValidIsoDate(s) Then NAVFITDecode = s: Exit Function
            If century <> 1900 And century <> 2000 Then Err.Raise vbObjectError + 5104, , "Choose 1900 or 2000 explicitly for the NAVFIT counseling date."
            If Len(s) <> 7 Then Err.Raise vbObjectError + 5104, , "Counseling text cannot be represented as a Swift date. Source value must be reviewed."
            If Not Left$(s, 2) Like "##" Or Not Right$(s, 2) Like "##" Then Err.Raise vbObjectError + 5104, , "Unsupported counseling date encoding."
            months = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
            For i = 0 To 11
                If UCase$(Mid$(s, 3, 3)) = months(i) Then m = i + 1
            Next
            If m = 0 Then Err.Raise vbObjectError + 5104, , "Invalid counseling month."
            y = century + CLng(Left$(s, 2)): dayValue = CLng(Right$(s, 2)): d = DateSerial(y, m, dayValue)
            If Year(d) <> y Or Month(d) <> m Or Day(d) <> dayValue Then Err.Raise vbObjectError + 5104, , "Invalid counseling date."
            NAVFITDecode = Format$(d, "yyyy-mm-dd")
        Case Else: Err.Raise vbObjectError + 5105, , "Unknown adapter conversion."
    End Select
End Function
Public Function NAVFITEncode(ByVal value As Variant, ByVal conversion As String) As Variant
    Dim i As Long, choices As Variant, d As Date, s As String, months As Variant
    If IsNull(value) Then
        If conversion = "flag" Or conversion = "observation" Then Err.Raise vbObjectError + 5107, , "Choose a checkbox value before export; native checkboxes cannot preserve an unknown value."
        NAVFITEncode = Null: Exit Function
    End If
    s = CStr(value)
    Select Case conversion
        Case "text": NAVFITEncode = s
        Case "date", "counseling"
            If conversion = "counseling" And IsCounselingStatus(s) Then NAVFITEncode = s: Exit Function
            If s = "" Then
                If conversion = "counseling" Then NAVFITEncode = "" Else NAVFITEncode = Null
                Exit Function
            End If
            If VarType(value) = vbDate Then
                If TimeValue(value) <> TimeSerial(0, 0, 0) Then Err.Raise vbObjectError + 5106, , "Export cannot silently discard a time from a date-only report field."
                s = Format$(value, "yyyy-mm-dd")
            End If
            If Not IsValidIsoDate(s) Then Err.Raise vbObjectError + 5106, , "Export requires a real ISO date."
            d = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Right$(s, 2)))
            If conversion = "date" Then
                NAVFITEncode = d
            Else
                If Year(d) < 1900 Or Year(d) > 2099 Then Err.Raise vbObjectError + 5106, , "NAVFIT counseling export supports explicitly resolved years from 1900 through 2099."
                months = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
                NAVFITEncode = Format$(d, "yy") & months(Month(d) - 1) & Format$(d, "dd")
            End If
        Case "flag", "observation"
            choices = Array("No", "Yes")
            If conversion = "observation" Then choices = Array("Observed", "NOB")
            If s = choices(0) Then NAVFITEncode = False: Exit Function
            If s = choices(1) Then NAVFITEncode = True: Exit Function
            Err.Raise vbObjectError + 5107, , "Choose a checkbox value before export. Blank is not silently treated as false."
        Case "trait"
            If s = "NOB" Then NAVFITEncode = CInt(0): Exit Function
            For i = 1 To 5
                If s = CStr(i) & ".0" Then NAVFITEncode = CInt(i): Exit Function
            Next
            Err.Raise vbObjectError + 5108, , "Unknown Swift trait value."
        Case "promotion"
            choices = Array("NOB", "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
            For i = 0 To 5
                If s = choices(i) Then NAVFITEncode = CInt(i): Exit Function
            Next
            Err.Raise vbObjectError + 5109, , "Choose a supported promotion recommendation."
        Case Else: Err.Raise vbObjectError + 5105, , "Unknown adapter conversion."
    End Select
End Function
Public Function NAVFITReadFields(ByVal source As DAO.Recordset, Optional ByVal century As Long = 0) As Object
    Dim result As Object, row As Variant, parts As Variant, key As String, value As Variant, kind As String
    Set result = CreateObject("Scripting.Dictionary"): kind = NAVFITKind(source!ReportType)
    result.Add "FormKind", kind
    For Each row In NAVFITMapRows(kind)
        parts = Split(row, "|"): key = parts(0)
        value = NAVFITDecode(source.Fields(parts(1)).Value, parts(2), century)
        If (key = "PeriodStart" Or key = "PeriodEnd") And Not IsNull(value) Then value = NAVFITEncode(value, "date")
        result.Add key, value
    Next
    Set NAVFITReadFields = result
End Function
Public Sub NAVFITWriteFields(ByVal destination As DAO.Recordset, ByVal values As Object)
    Dim row As Variant, parts As Variant, value As Variant, f As DAO.Field
    Dim fieldKey As String, failure As Long, message As String
    On Error GoTo Failed
    destination!ReportType = NAVFITType(values("FormKind"))
    For Each row In NAVFITMapRows(values("FormKind"))
        parts = Split(row, "|"): fieldKey = parts(0)
        value = NAVFITEncode(values(fieldKey), parts(2)): Set f = destination.Fields(parts(1))
        If Not IsNull(value) And f.Type = dbText Then
            If Len(CStr(value)) > f.Size Then Err.Raise vbObjectError + 5110, , parts(0) & " exceeds native capacity; no truncation is permitted."
        End If
        f.Value = value
    Next
    destination!IsValidated = False
    Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    If fieldKey <> "" Then message = UXFieldCaption(fieldKey, values("FormKind")) & ": " & message
    Err.Raise failure, "NAVFITWriteFields", message
End Sub
' Three-way merge: unchanged returned values preserve Swift edits; distinct edits conflict.
Public Function NAVFITMergeDecision(ByVal baseline As Variant, ByVal current As Variant, ByVal incoming As Variant) As String
    If SameFieldValue(incoming, baseline) Or SameFieldValue(incoming, current) Then
        NAVFITMergeDecision = "Keep current"
    ElseIf SameFieldValue(current, baseline) Then
        NAVFITMergeDecision = "Accept returned"
    Else
        NAVFITMergeDecision = "Conflict"
    End If
End Function
' Normal commit and revision services remain the only write path for mapped values.
Public Function NAVFITCommitFields(ByVal reportID As Long, ByVal values As Object, ByVal member As String, ByVal reason As String) As Long
    NAVFITCommitFields = CommitDraft(reportID, values("FormKind"), "Imported " & values("FormKind"), member, values("PeriodStart"), values("PeriodEnd"), "", reason, values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
End Function
