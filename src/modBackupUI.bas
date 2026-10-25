Attribute VB_Name = "modBackupUI"
Option Compare Binary
Option Explicit
'Backups are verified data files written inside Access: no helper process, so
'endpoint rules that block Office child processes do not stop them. Restore is
'the Settings & Help import into a new blank workspace, which reads any backup.
Public Function GuidedWorkspaceBackup() As Boolean
    Dim folder As Object, destination As String, evidence As String
    On Error GoTo Failed
    Set folder = Application.FileDialog(4): folder.Title = "Choose a folder for the new verified backup file"
    If folder.Show <> -1 Then Exit Function
    destination = folder.SelectedItems(1) & "\Swift_NAVFIT_backup_" & Format$(Now(), "yyyymmdd_hhnnss") & "_" & Replace(Replace(NewIdentity(), "{", ""), "}", "") & ".accdb"
    Forms!frmBackup!lblStatus.Caption = "Writing and verifying the backup. Keep Access open."
    DoEvents
    evidence = WriteWorkspaceBackup(destination)
    If CurrentProject.AllForms("frmBackup").IsLoaded Then
        Forms!frmBackup!txtResult = evidence
        Forms!frmBackup!lblStatus.Caption = "Backup verified: every row and field matches. Saved " & destination
    End If
    GuidedWorkspaceBackup = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmBackup").IsLoaded Then Forms!frmBackup!lblStatus.Caption = Err.Description
End Function

'Writes every local table, keys and indexes included, into a new .accdb, then
'checks every row and field against this workspace with the importer's own
'verifier. Returns the verification record, also saved beside the backup.
Public Function WriteWorkspaceBackup(ByVal destination As String) As String
    Dim db As DAO.Database, backup As DAO.Database, table As DAO.TableDef, q As DAO.QueryDef, names As Collection, name As Variant
    Dim prefix As String, counts As String, rows As Long, total As Long, created As Boolean, evidence As String, operation As Long
    Dim failure As Long, message As String, source As String, folder As String
    On Error GoTo Failed
    If LCase$(Right$(destination, 6)) <> ".accdb" Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "A backup is a new .accdb data file."
    folder = Left$(destination, InStrRev(destination, "\") - 1)
    If Len(folder) = 0 Or Len(Dir$(folder, vbDirectory)) = 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination folder does not exist."
    If Len(Dir$(destination)) > 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination already exists. A backup always creates a new file."
    If Len(Dir$(destination & ".verification.json")) > 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination verification file already exists; it was not overwritten."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Not UXCloseEditorForBackup() Then Err.Raise vbObjectError + 4801, "modBackupUI.WriteWorkspaceBackup", "Pending editor values could not be checkpointed and closed. Save or close the editor before retrying."
    End If
    Set db = CurrentDb: Set names = New Collection
    For Each table In db.TableDefs
        If Left$(table.Name, 4) <> "MSys" And Left$(table.Name, 1) <> "~" Then
            If Len(table.Connect) > 0 Then Err.Raise vbObjectError + 4802, "modBackupUI.WriteWorkspaceBackup", "Linked tables cannot be backed up as one file."
            names.Add table.Name
        End If
    Next
    Set backup = DBEngine.CreateDatabase(destination, dbLangGeneral, dbVersion120): created = True
    backup.Close: Set backup = Nothing
    For Each name In names
        DoCmd.TransferDatabase acExport, "Microsoft Access", destination, acTable, CStr(name), CStr(name), False
    Next
    prefix = "zSwiftBackup" & Left$(Replace(Replace(Replace(NewIdentity(), "{", ""), "}", ""), "-", ""), 20)
    Set backup = DBEngine.OpenDatabase(destination, False, True)
    For Each name In names
        Set q = db.CreateQueryDef(prefix & name, "SELECT * FROM [" & name & "] IN " & Chr$(34) & destination & Chr$(34)): q.Close: Set q = Nothing
        rows = VerifyMigrationTable(db, backup.TableDefs(CStr(name)), prefix & name)
        If Len(counts) > 0 Then counts = counts & ","
        counts = counts & JsonString(CStr(name)) & ":" & rows
        total = total + rows
    Next
    backup.Close: Set backup = Nothing
    RemoveBackupQueries prefix
    source = CurrentProject.FullName
    evidence = "{""Status"":""PASS"",""Mode"":""Data backup"",""Source"":" & JsonString(source) & ",""Destination"":" & JsonString(destination) & _
        ",""SHA256"":""" & SourceFileHash(destination) & """,""Schema"":" & SCHEMA_VERSION & ",""AppVersion"":" & JsonString(APP_VERSION) & _
        ",""RowsVerified"":" & total & ",""CreatedAt"":""" & Format$(Now(), "yyyy-mm-dd\Thh:nn:ss") & """,""Counts"":{" & counts & "}}"
    WriteNewTextFile destination & ".verification.json", evidence
    operation = BeginOperation("Data backup", destination)
    FinishOperation operation, "Completed", "Verified " & total & " rows. Restore with Settings & Help > Import data into a new blank workspace."
    WriteWorkspaceBackup = evidence
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not q Is Nothing Then q.Close
    If Not backup Is Nothing Then backup.Close
    RemoveBackupQueries prefix
    'Only the incomplete file this call created is removed; nothing else is touched.
    If created Then Kill destination
    On Error GoTo 0
    Err.Raise failure, "modBackupUI.WriteWorkspaceBackup", message
End Function
Private Sub RemoveBackupQueries(ByVal prefix As String)
    Dim db As DAO.Database, i As Long
    If Len(prefix) = 0 Then Exit Sub
    Set db = CurrentDb
    For i = db.QueryDefs.Count - 1 To 0 Step -1
        If Left$(db.QueryDefs(i).Name, Len(prefix)) = prefix Then db.QueryDefs.Delete db.QueryDefs(i).Name
    Next
End Sub
Private Function JsonString(ByVal value As String) As String
    Dim i As Long, c As String, result As String
    For i = 1 To Len(value)
        c = Mid$(value, i, 1)
        Select Case AscW(c)
            Case 34: result = result & "\"""
            Case 92: result = result & "\\"
            Case 0 To 31: result = result & "\u" & Right$("000" & Hex$(AscW(c)), 4)
            Case Else: result = result & c
        End Select
    Next
    JsonString = """" & result & """"
End Function
Private Sub WriteNewTextFile(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text
    stream.SaveToFile path, 1
    stream.Close
End Sub
Public Function ShowCopyVerification() As Boolean
    Dim picker As Object, stream As Object, path As String
    On Error GoTo Failed
    Set picker = Application.FileDialog(3): picker.Title = "Choose the backup verification file": picker.AllowMultiSelect = False
    picker.Filters.Clear: picker.Filters.Add "Verification result", "*.verification.json"
    If picker.Show <> -1 Then Exit Function
    path = picker.SelectedItems(1)
    If FileLen(path) > 500000 Then Err.Raise vbObjectError + 4802, "modBackupUI.ShowCopyVerification", "Verification result is too large."
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile path
    Forms!frmBackup!txtResult = stream.ReadText: stream.Close
    ShowCopyVerification = True: Exit Function
Failed:
    Forms!frmBackup!lblStatus.Caption = Err.Description
End Function
