Attribute VB_Name = "modSchemaUpgrade"
Option Compare Database
Option Explicit
' Automatic schema upgrade of a separate data file (backlog S01, owner decision 13 Nov
' 2026). A newer Swift NAVFIT pointed at an older data file copies the file beside
' itself, checks the copy, applies the schema steps between the two versions in order,
' and stamps each version as it goes. If any step fails, the data file is put back from
' the copy, so the user is left with the file they had.
'
' Every schema change adds its step below, written so that a data file of the previous
' version plus the step has exactly the structure a fresh build of the new version
' has. test-data-link checks that against the captured definition of the previous
' schema (tests/historical-schemas).

' Oldest data file schema these steps start from.
Public Const UPGRADE_FLOOR As Long = 47

' The statements that take a data file from toVersion - 1 to toVersion.
Private Function SchemaStepSQL(ByVal toVersion As Long) As Variant
    Select Case toVersion
        Case 48
            'Content checks that can be turned off.
            SchemaStepSQL = Array("CREATE TABLE RestrictedCheckOff (GroupID TEXT(60) CONSTRAINT PK_RestrictedCheckOff PRIMARY KEY)")
        Case Else
            Err.Raise vbObjectError + 9820, "modSchemaUpgrade.SchemaStepSQL", "No schema step leads to schema " & toVersion & "."
    End Select
End Function

' The data file's schema version; 0 when it has none.
Public Function DataFileSchema(ByVal filePath As String) As Long
    Dim db As DAO.Database, r As DAO.Recordset
    Set db = DBEngine.Workspaces(0).OpenDatabase(filePath, False, True)
    On Error GoTo Done
    Set r = db.OpenRecordset("SELECT SchemaVersion FROM AppMetadata", dbOpenSnapshot)
    If Not r.EOF Then DataFileSchema = CLng(Nz(r!SchemaVersion, 0))
    r.Close
Done:
    db.Close
End Function

' True when this version can upgrade the data file in place.
Public Function DataFileUpgradable(ByVal filePath As String) As Boolean
    Dim version As Long
    On Error GoTo Done
    If Not FileExists(filePath) Then Exit Function
    version = DataFileSchema(filePath)
    DataFileUpgradable = (version >= UPGRADE_FLOOR And version < SCHEMA_VERSION)
Done:
End Function

' Field names per table, to find what a step added.
Private Function FieldSnapshot(ByVal db As DAO.Database) As Object
    Dim t As DAO.TableDef, f As DAO.Field, seen As Object
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    db.TableDefs.Refresh
    For Each t In db.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" Then
            For Each f In t.Fields
                seen(t.Name & "." & f.Name) = True
            Next
        End If
    Next
    Set FieldSnapshot = seen
End Function

' Upgrades the data file to this version's schema and returns the path of the copy made
' before it. Refuses a newer file, one older than the steps, and one another session has
' open.
Public Function UpgradeDataFile(ByVal filePath As String) As String
    Dim version As Long, target As Long, backup As String, db As DAO.Database, statement As Variant, before As Object
    Dim t As DAO.TableDef, f As DAO.Field, stage As String, failure As Long, message As String, changed As Boolean, copied As Boolean, fso As Object
    On Error GoTo Failed
    stage = "Check data file"
    If Not FileExists(filePath) Then Err.Raise vbObjectError + 9821, "modSchemaUpgrade.UpgradeDataFile", "The data file was not found: " & filePath
    version = DataFileSchema(filePath)
    If version = SCHEMA_VERSION Then Exit Function
    If version > SCHEMA_VERSION Then Err.Raise vbObjectError + 9822, "modSchemaUpgrade.UpgradeDataFile", "The data file is schema " & version & ", newer than this version of Swift NAVFIT (schema " & SCHEMA_VERSION & "). Use a newer Swift NAVFIT."
    If version < UPGRADE_FLOOR Then Err.Raise vbObjectError + 9823, "modSchemaUpgrade.UpgradeDataFile", "The data file is schema " & version & ", older than this version can upgrade (schema " & UPGRADE_FLOOR & " onward)."
    'This session's own handle on the data file would block the exclusive open.
    SwiftDataRelease
    stage = "Check no one else has it open"
    Set db = DBEngine.Workspaces(0).OpenDatabase(filePath, True)
    db.Close: Set db = Nothing
    stage = "Copy the data file"
    backup = Left$(filePath, Len(filePath) - 6) & " before schema " & SCHEMA_VERSION & " " & Format$(Now(), "yyyymmdd-hhnnss") & ".accdb"
    If Len(Dir$(backup)) > 0 Then Err.Raise vbObjectError + 9824, "modSchemaUpgrade.UpgradeDataFile", "A copy with this name already exists: " & backup
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile filePath, backup, False: copied = True
    If SourceFileHash(backup) <> SourceFileHash(filePath) Then Err.Raise vbObjectError + 9825, "modSchemaUpgrade.UpgradeDataFile", "The copy of the data file does not match it. Nothing was changed."
    stage = "Open the data file"
    Set db = DBEngine.Workspaces(0).OpenDatabase(filePath, True)
    For target = version + 1 To SCHEMA_VERSION
        stage = "Schema step to " & target
        Set before = FieldSnapshot(db)
        changed = True
        For Each statement In SchemaStepSQL(target)
            db.Execute CStr(statement), dbFailOnError
        Next
        'Builds let optional text be empty; give what the step added the same rule.
        db.TableDefs.Refresh
        For Each t In db.TableDefs
            If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" Then
                For Each f In t.Fields
                    If Not before.Exists(t.Name & "." & f.Name) Then
                        If (f.Type = dbText Or f.Type = dbMemo) And Not f.Required Then f.AllowZeroLength = True
                    End If
                Next
            End If
        Next
        db.Execute "UPDATE AppMetadata SET SchemaVersion=" & target & ", AppVersion='" & Replace(APP_VERSION, "'", "''") & "'", dbFailOnError
    Next
    db.Close: Set db = Nothing
    stage = "Check the upgraded data file"
    If DataFileSchema(filePath) <> SCHEMA_VERSION Then Err.Raise vbObjectError + 9826, "modSchemaUpgrade.UpgradeDataFile", "The data file did not reach schema " & SCHEMA_VERSION & "."
    UpgradeDataFile = backup
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then db.Close: Set db = Nothing
    If changed Then
        'Put back the file the user had.
        Kill filePath
        fso.CopyFile backup, filePath, False
        If SourceFileHash(filePath) = SourceFileHash(backup) Then
            message = message & " The data file was put back as it was before the upgrade."
        Else
            message = message & " The data file could not be put back; the copy from before the upgrade is " & backup & "."
        End If
    ElseIf copied Then
        'Nothing changed: the copy this call made is not needed.
        Kill backup
    End If
    On Error GoTo 0
    If stage = "Check no one else has it open" Then message = "The data file is open in another Swift NAVFIT session. Close it there, then try again. (" & message & ")"
    Err.Raise failure, "modSchemaUpgrade.UpgradeDataFile", stage & ": " & message
End Function

' Upgrades when needed and tells the user where the copy is. "" when the data file is
' ready or needs nothing; otherwise why the upgrade did not happen.
Public Function UpgradeDataFileIfNeeded(ByVal filePath As String) As String
    Dim version As Long, backup As String
    On Error GoTo Failed
    If Not FileExists(filePath) Then Exit Function
    version = DataFileSchema(filePath)
    If version = SCHEMA_VERSION Or version = 0 Then Exit Function
    backup = UpgradeDataFile(filePath)
    If Application.UserControl And Len(backup) > 0 Then
        MsgBox "Swift NAVFIT upgraded your data file from schema " & version & " to schema " & SCHEMA_VERSION & "." & vbCrLf & vbCrLf & "A copy of the file from before the upgrade is kept beside it:" & vbCrLf & backup, vbInformation, "Swift NAVFIT data file"
    End If
    Exit Function
Failed:
    UpgradeDataFileIfNeeded = Err.Description
End Function

' True when the file exists; a missing drive or share is not an error here.
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = Len(Dir$(filePath)) > 0
End Function
