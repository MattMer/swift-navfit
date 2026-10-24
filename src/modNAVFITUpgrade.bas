Attribute VB_Name = "modNAVFITUpgrade"
Option Compare Binary
Option Explicit

' Exact additive changes and NULL-to-empty initialization from the supplied v1.0
' NAVFIT98A v33 Patch 2 conversion utility. No source application code is executed.
Public Function NAVFITPatchColumns() As Variant
    NAVFITPatchColumns = Array("Reports|DoDID|10", "Reports|RSDoDID|10", _
        "Folders|RSAddress1|30", "Folders|RSAddress2|30", "Folders|RSCity|15", _
        "Folders|RSState|2", "Folders|RSZipCd|9", "Folders|RSPhone|10", "Folders|RSDSN|7")
End Function
' Return False only for a missing field; propagate unrelated DAO failures.
Public Function NAVFITHasField(ByVal table As DAO.TableDef, ByVal name As String) As Boolean
    Dim f As DAO.Field
    For Each f In table.Fields
        If StrComp(f.Name, name, vbBinaryCompare) = 0 Then NAVFITHasField = True: Exit Function
    Next
End Function
Private Function IsPatchColumn(ByVal table As String, ByVal name As String) As Boolean
    Dim entry As Variant, parts As Variant
    For Each entry In NAVFITPatchColumns()
        parts = Split(entry, "|")
        If table = parts(0) And name = parts(1) Then IsPatchColumn = True: Exit Function
    Next
End Function
' Validate every known field by name/type/size, allowing only the nine documented
' omissions on native v33 input. Upgrader-appended columns need not retain ordinals.
Public Sub ValidateNAVFITSourceSchema(ByVal db As DAO.Database, ByVal prefix As String, Optional ByVal allowV33 As Boolean = False)
    Dim table As Variant, entry As Variant, parts As Variant, t As DAO.TableDef, f As DAO.Field, expected As Long
    For Each table In Array("Folders", "Reports", "Summary")
        Set t = db.TableDefs(prefix & table)
        If t.Attributes <> 0 Then Err.Raise vbObjectError + 404, "modNAVFITUpgrade.ValidateNAVFITSourceSchema", "Linked or hidden source tables are unsupported."
        expected = 0
        For Each entry In NAVFITSchemaFields(CStr(table))
            parts = Split(entry, "|")
            If NAVFITHasField(t, CStr(parts(0))) Then
                Set f = t.Fields(parts(0)): expected = expected + 1
                If f.Type <> CLng(parts(1)) Or f.Size <> CLng(parts(2)) Then Err.Raise vbObjectError + 403, "modNAVFITUpgrade.ValidateNAVFITSourceSchema", "Unsupported NAVFIT field: " & table & "." & parts(0)
            ElseIf Not (allowV33 And prefix = "" And IsPatchColumn(CStr(table), CStr(parts(0)))) Then
                Err.Raise vbObjectError + 403, "modNAVFITUpgrade.ValidateNAVFITSourceSchema", "Missing NAVFIT field: " & table & "." & parts(0)
            End If
        Next
        If t.Fields.Count <> expected Then Err.Raise vbObjectError + 402, "modNAVFITUpgrade.ValidateNAVFITSourceSchema", "Unexpected NAVFIT fields in " & table
    Next
End Sub
' Detect missing Patch 2 fields only after complete base-schema validation.
Public Function NAVFITNeedsV33Upgrade(ByVal db As DAO.Database) As Boolean
    Dim entry As Variant, parts As Variant
    For Each entry In NAVFITPatchColumns()
        parts = Split(entry, "|")
        If Not NAVFITHasField(db.TableDefs(parts(0)), CStr(parts(1))) Then NAVFITNeedsV33Upgrade = True: Exit Function
    Next
End Function
' Apply to an owned disposable copy in its separate DAO workspace. A failed DDL
' or initialization step rolls back and the caller removes the disposable file.
Public Sub UpgradeNAVFITV33Copy(ByVal path As String, ByVal reader As DAO.Workspace)
    Dim db As DAO.Database, entry As Variant, parts As Variant, field As DAO.Field
    Dim trans As Boolean, failure As Long, message As String
    On Error GoTo Failed
    Set db = reader.OpenDatabase(path, True)
    ValidateNAVFITSourceSchema db, "", True
    reader.BeginTrans: trans = True
    For Each entry In NAVFITPatchColumns()
        parts = Split(entry, "|")
        If Not NAVFITHasField(db.TableDefs(parts(0)), CStr(parts(1))) Then
            Set field = db.TableDefs(parts(0)).CreateField(parts(1), dbText, CLng(parts(2)))
            field.AllowZeroLength = True: db.TableDefs(parts(0)).Fields.Append field
        End If
        db.TableDefs(parts(0)).Fields(parts(1)).AllowZeroLength = True
        db.Execute "UPDATE [" & parts(0) & "] SET [" & parts(1) & "] = " & Chr$(39) & Chr$(39) & " WHERE [" & parts(1) & "] IS NULL", dbFailOnError
    Next
    ValidateNAVFITSourceSchema db, ""
    reader.CommitTrans: trans = False: db.Close
    Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If trans Then reader.Rollback
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Err.Raise failure, "UpgradeNAVFITV33Copy", message
End Sub
Private Function NAVFITSchemaFields(ByVal table As String) As Variant
    Select Case table
        Case "Folders"
            Dim spec As String
            spec = "FolderName|10|50;FolderID|4|4;Parent|3|2;FullName|10|27;Rate|10|5;Desig|10|12;SSN|10|9;Active|1|1"
            spec = spec & ";TAR|1|1;Inactive|1|1;ATADSW|1|1;UIC|10|5;ShipStation|10|18;PromotionStatus|10|8;DateReported|8|8;Periodic|1|1"
            spec = spec & ";DetInd|1|1;Frocking|1|1;Special|1|1;FromDate|8|8;ToDate|8|8;NOB|1|1;Regular|1|1;Concurrent|1|1"
            spec = spec & ";OpsCdr|1|1;PhysicalReadiness|10|4;BilletSubcat|10|10;ReportingSenior|10|18;RSGrade|10|5;RSDesig|10|5;RSTitle|10|14;RSUIC|10|5"
            spec = spec & ";RSSSN|10|9;Achievements|12|0;PrimaryDuty|10|14;Duties|12|0;DateCounseled|10|8;Counseler|10|20;PROF|3|2;QUAL|3|2"
            spec = spec & ";EO|3|2;MIL|3|2;PA|3|2;TEAM|3|2;LEAD|3|2;MIS|3|2;TAC|3|2;RecommendA|10|20"
            spec = spec & ";RecommendB|10|20;Rater|10|18;RaterDate|8|8;CommentsPitch|10|8;Comments|12|0;Qualifications|12|0;PromotionRecom|3|2;SummarySP|10|3"
            spec = spec & ";SummaryProg|10|3;SummaryProm|10|3;SummaryMP|10|3;SummaryEP|10|3;RetentionYes|1|1;RetentionNo|1|1;RSAddress|12|0;RSAddress1|10|30"
            spec = spec & ";RSAddress2|10|30;RSCity|10|15;RSState|10|2;RSZipCd|10|9;RSPhone|10|10;RSDSN|10|7;SeniorRater|10|18;SeniorRaterDate|8|8"
            spec = spec & ";StatementYes|1|1;StatementNo|1|1;RSInfo|12|0;AutoSummary|1|1;UserComments|12|0;Psswrd|10|8"
        Case "Reports"
            spec = "Parent|10|12;ReportID|4|4;ReportType|10|15;FullName|10|27;FirstName|10|27;MI|10|3;LastName|10|27;Suffix|10|10"
            spec = spec & ";Rate|10|5;Desig|10|12;SSN|10|9;DoDID|10|10;Active|1|1;TAR|1|1;Inactive|1|1;ATADSW|1|1"
            spec = spec & ";UIC|10|5;ShipStation|10|18;PromotionStatus|10|8;DateReported|8|8;Periodic|1|1;DetInd|1|1;Frocking|1|1;Special|1|1"
            spec = spec & ";FromDate|8|8;ToDate|8|8;NOB|1|1;Regular|1|1;Concurrent|1|1;OpsCdr|1|1;PhysicalReadiness|10|4;PhysicalReadiness2|10|15"
            spec = spec & ";PhysicalReadinessDt|8|8;BilletSubcat|10|10;RSLastName|10|18;RSFI|10|1;RSMI|10|3;ReportingSenior|10|18;RSGrade|10|5;RSDesig|10|5"
            spec = spec & ";RSTitle|10|14;RSUIC|10|5;RSSSN|10|9;RSDoDID|10|10;Achievements|12|0;PrimaryDuty|10|14;Duties|12|0;DateCounseled|10|8"
            spec = spec & ";Counseler|10|20;CounselerLN|10|20;CounselerFI|10|1;CounselerMI|10|3;PROF|3|2;PROFDN1|10|255;PROFDN2|10|255;PROFDN3|10|255"
            spec = spec & ";QUAL|3|2;QUALDN1|10|255;QUALDN2|10|255;QUALDN3|10|255;EO|3|2;EODN1|10|255;EODN2|10|255;EODN3|10|255"
            spec = spec & ";MIL|3|2;MILDN1|10|255;MILDN2|10|255;MILDN3|10|255;PA|3|2;PADN1|10|255;PADN2|10|255;PADN3|10|255"
            spec = spec & ";TEAM|3|2;TEAMDN1|10|255;TEAMDN2|10|255;TEAMDN3|10|255;LEAD|3|2;LEADDN1|10|255;LEADDN2|10|255;LEADDN3|10|255"
            spec = spec & ";MIS|3|2;MISDN1|10|255;MISDN2|10|255;MISDN3|10|255;TAC|3|2;TACDN1|10|255;TACDN2|10|255;TACDN3|10|255"
            spec = spec & ";RecommendA|10|20;RecommendB|10|20;Rater|10|28;RaterDate|8|8;CommentsPitch|10|8;Comments|12|0;Qualifications|12|0;PromotionRecom|3|2"
            spec = spec & ";SummaryRank|4|4;SummarySP|10|3;SummaryProg|10|3;SummaryProm|10|3;SummaryMP|10|3;SummaryEP|10|3;RetentionYes|1|1;RetentionNo|1|1"
            spec = spec & ";RSCA|20|16;RSAddress|12|0;RSAddress1|10|30;RSAddress2|10|30;RSCity|10|15;RSState|10|2;RSZipCd|10|9;RSPhone|10|10"
            spec = spec & ";RSDSN|10|7;SeniorRater|10|28;SeniorRaterDate|8|8;StatementYes|1|1;StatementNo|1|1;RSInfo|12|0;RRSFI|10|1;RRSMI|10|3"
            spec = spec & ";RRSLastName|10|25;RRSGrade|10|5;RRSCommand|10|30;RRSUIC|10|5;UserComments|12|0;Psswrd|10|8;Standards|10|30;IsValidated|1|1"
        Case "Summary"
            spec = "ProfileName|10|50;OnTimeYes|1|1;OnTimeNo|1|1;ISICAddress|12|0;EMail|10|50;Phone|10|10"
    End Select
    NAVFITSchemaFields = Split(spec, ";")
End Function
