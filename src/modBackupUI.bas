Attribute VB_Name = "modBackupUI"
Option Compare Binary
Option Explicit
'Backups are verified data files written inside Access: no helper process, so
'endpoint rules that block Office child processes do not stop them. Restore is
'the Settings & Help import into a new blank workspace, which reads any backup.
'Access stops a database at 2 GB; Settings & Help warns from 1.5 GB (L03).
Private Const WorkspaceLimitBytes As Double = 2147483648#
Private Const WorkspaceWarnBytes As Double = 1610612736#
Public Function WorkspaceFileBytes() As Double
    'FileLen reports an open file's size from before it was opened.
    WorkspaceFileBytes = CreateObject("Scripting.FileSystemObject").GetFile(SwiftDataPath()).Size
End Function
Public Function WorkspaceSizeNearLimit(ByVal bytes As Double) As Boolean
    WorkspaceSizeNearLimit = (bytes >= WorkspaceWarnBytes)
End Function
Public Function WorkspaceSizeText(ByVal bytes As Double) As String
    If WorkspaceSizeNearLimit(bytes) Then
        WorkspaceSizeText = "Warning: this workspace file is " & Format$(bytes / 1073741824#, "0.00") & " GB of the 2 GB Access limit. Back it up and start a new workspace soon; see the workspace guide."
    Else
        WorkspaceSizeText = "Workspace file: " & Format$(bytes / 1048576#, "0.0") & " MB of the 2 GB Access limit (" & Format$(bytes / WorkspaceLimitBytes, "0%") & ")."
    End If
End Function
Public Function UXSettingsActivate() As Boolean
    Dim bytes As Double
    UXSettingsActivate = UXLoadContext()
    bytes = WorkspaceFileBytes()
    With Forms!frmUXSettings!uxWorkspaceSize
        .Caption = WorkspaceSizeText(bytes): .FontBold = WorkspaceSizeNearLimit(bytes)
    End With
    If UXHasControl(Forms!frmUXSettings, "txtAbout") Then Forms!frmUXSettings!txtAbout = AppAboutText(APP_RELEASE_DATE)
End Function
'Version, release date and where to find newer releases. Text only: no network use.
Public Function AppAboutText(ByVal releaseDate As String) As String
    Dim text As String, released As Variant
    text = "Swift NAVFIT " & APP_VERSION & ", schema " & SCHEMA_VERSION & "." & vbCrLf
    released = ReleaseDateValue(releaseDate)
    If IsNull(released) Then
        text = text & "Development build, not a published release." & vbCrLf
    Else
        text = text & "Released " & Format$(released, "d mmmm yyyy") & "." & vbCrLf
    End If
    AppAboutText = text & "Newer releases: " & APP_UPDATES_URL & vbCrLf & "Copy the address into a web browser. Swift NAVFIT does not connect to the internet."
End Function
'A one-line Overview note once a release is 90 or more days old; empty otherwise.
Public Function UpdateReminderText(ByVal releaseDate As String, ByVal today As Date) As String
    Dim released As Variant
    released = ReleaseDateValue(releaseDate)
    If IsNull(released) Then Exit Function
    If DateDiff("d", released, today) < 90 Then Exit Function
    UpdateReminderText = "This version is from " & Format$(released, "d mmm yyyy") & ", over 90 days ago. A newer release may be available: see Settings & Help, About this version."
End Function
'yyyy-mm-dd to a date without relying on regional settings; Null when empty or invalid.
Private Function ReleaseDateValue(ByVal releaseDate As String) As Variant
    ReleaseDateValue = Null
    If Not releaseDate Like "####-##-##" Then Exit Function
    On Error Resume Next
    ReleaseDateValue = DateSerial(CInt(Left$(releaseDate, 4)), CInt(Mid$(releaseDate, 6, 2)), CInt(Right$(releaseDate, 2)))
End Function
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
    If CurrentProject.AllForms("frmBackup").IsLoaded Then Forms!frmBackup!lblStatus.Caption = SwiftErrorText()
End Function

'Writes every table, keys and indexes included, into a new .accdb (with a separate
'data file, a copy of that file), then
'checks every row and field against this workspace with the importer's own
'verifier. Returns the verification record, also saved beside the backup.
Public Function WriteWorkspaceBackup(ByVal destination As String) As String
    Dim db As DAO.Database, table As DAO.TableDef, names As Collection, attempt As Long
    Dim counts As String, total As Long, created As Boolean, evidence As String, operation As Long
    Dim failure As Long, message As String, source As String, folder As String, receipt As Boolean, published As Boolean
    Dim linked As Boolean, dataFile As String
    On Error GoTo Failed
    If LCase$(Right$(destination, 6)) <> ".accdb" Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "A backup is a new .accdb data file."
    folder = Left$(destination, InStrRev(destination, "\") - 1)
    If Len(folder) = 0 Or Len(Dir$(folder, vbDirectory)) = 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination folder does not exist."
    If Len(Dir$(destination)) > 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination already exists. A backup always creates a new file."
    If Len(Dir$(destination & ".verification.json")) > 0 Then Err.Raise vbObjectError + 4800, "modBackupUI.WriteWorkspaceBackup", "Destination verification file already exists; it was not overwritten."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Not UXCloseEditorForBackup() Then Err.Raise vbObjectError + 4801, "modBackupUI.WriteWorkspaceBackup", "Pending editor values could not be checkpointed and closed. Save or close the editor before retrying."
    End If
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4830, "modBackupUI.WriteWorkspaceBackup", "Save or discard the unsaved Members-grid cells first. A backup copies only saved values."
    Set db = CurrentDb: Set names = New Collection
    linked = SwiftDataIsLinked(): dataFile = SwiftDataPath()
    For Each table In db.TableDefs
        If SwiftWorkspaceTable(table.Name) Then
            'Either every table is in this file, or every table is in the one data file.
            If (Len(table.Connect) > 0) <> linked Then Err.Raise vbObjectError + 4802, "modBackupUI.WriteWorkspaceBackup", "Linked tables cannot be backed up as one file."
            If linked Then
                If StrComp(Mid$(table.Connect, InStr(1, table.Connect, "DATABASE=", vbTextCompare) + 9), dataFile, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 4802, "modBackupUI.WriteWorkspaceBackup", "Linked tables cannot be backed up as one file: they are in more than one file."
            End If
            names.Add table.Name
        End If
    Next
    'A shared data file can change while it is copied, when others save at that moment;
    'the check then differs and the copy is made again, up to three times.
    For attempt = 1 To IIf(linked, 3, 1)
        If attempt > 1 Then SwiftPause 2000 * attempt
        On Error Resume Next
        BackupCopy db, names, destination, linked, counts, total
        failure = Err.Number: message = Err.Description
        On Error GoTo Failed
        If failure = 0 Then Exit For
        If Not (linked And BackupCopyBusy(failure)) Then Err.Raise failure, "modBackupUI.WriteWorkspaceBackup", message
    Next
    If failure <> 0 Then Err.Raise vbObjectError + 4831, "modBackupUI.WriteWorkspaceBackup", "Others were saving while the backup was copied, so it could not be verified. No backup was kept. Try again in a minute."
    created = True
    source = SwiftDataPath()
    evidence = "{""Status"":""PASS"",""Mode"":""Data backup"",""Source"":" & JsonString(source) & ",""Destination"":" & JsonString(destination) & _
        ",""SHA256"":""" & SourceFileHash(destination) & """,""Schema"":" & SCHEMA_VERSION & ",""AppVersion"":" & JsonString(APP_VERSION) & _
        ",""RowsVerified"":" & total & ",""CreatedAt"":""" & Format$(Now(), "yyyy-mm-dd\Thh:nn:ss") & """,""Counts"":{" & counts & "}}"
    receipt = True: WriteNewTextFile destination & ".verification.json", evidence: published = True
    operation = BeginOperation("Data backup", destination)
    FinishOperation operation, "Completed", "Verified " & total & " rows. Restore with Settings & Help > Import data into a new blank workspace."
    WriteWorkspaceBackup = evidence
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    'A verified backup with its receipt is kept when only the journal entry fails.
    'Otherwise only the incomplete files this call created are removed.
    If Not published Then
        If created Then Kill destination
        If receipt Then Kill destination & ".verification.json"
    End If
    On Error GoTo 0
    If published Then message = "The backup was written and verified, and its verification file was saved: " & destination & ". Only the local backup journal entry failed: " & message
    Err.Raise failure, "modBackupUI.WriteWorkspaceBackup", message
End Function

'Makes the backup file and checks every row and field against the open workspace.
'On failure, removes the file and the queries it made, then raises the error.
Private Sub BackupCopy(ByVal db As DAO.Database, ByVal names As Collection, ByVal destination As String, ByVal linked As Boolean, ByRef counts As String, ByRef total As Long)
    Dim backup As DAO.Database, q As DAO.QueryDef, name As Variant, rows As Long, prefix As String, created As Boolean, failure As Long, message As String
    On Error GoTo Failed
    counts = "": total = 0
    If linked Then
        'The data file holds only tables, so the backup is a copy of it. The check
        'below still compares every row and field with the open workspace.
        SwiftFreshRead
        CreateObject("Scripting.FileSystemObject").CopyFile SwiftDataPath(), destination, False: created = True
    Else
        Set backup = DBEngine.CreateDatabase(destination, dbLangGeneral, dbVersion120): created = True
        backup.Close: Set backup = Nothing
        For Each name In names
            DoCmd.TransferDatabase acExport, "Microsoft Access", destination, acTable, CStr(name), CStr(name), False
        Next
    End If
    prefix = "zSwiftBackup" & Left$(Replace(Replace(Replace(NewIdentity(), "{", ""), "}", ""), "-", ""), 20)
    Set backup = DBEngine.OpenDatabase(destination, False, True)
    For Each name In names
        'Who has the file open changes by the second and is not workspace data: copied, not checked.
        If SwiftTransientTable(CStr(name)) Then GoTo NextName
        Set q = db.CreateQueryDef(prefix & name, "SELECT * FROM [" & name & "] IN " & Chr$(34) & destination & Chr$(34)): q.Close: Set q = Nothing
        rows = VerifyMigrationTable(db, backup.TableDefs(CStr(name)), prefix & name)
        If Len(counts) > 0 Then counts = counts & ","
        counts = counts & JsonString(CStr(name)) & ":" & rows
        total = total + rows
NextName:
    Next
    backup.Close: Set backup = Nothing
    RemoveBackupQueries prefix
    Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not q Is Nothing Then q.Close
    If Not backup Is Nothing Then backup.Close
    RemoveBackupQueries prefix
    'The destination did not exist before this backup began.
    If created Or Len(Dir$(destination)) > 0 Then Kill destination
    On Error GoTo 0
    Err.Raise failure, "modBackupUI.BackupCopy", message
End Sub

'A copy that differs from the data file, or that could not be read or made, because
'someone was saving at that moment.
Private Function BackupCopyBusy(ByVal failure As Long) As Boolean
    Select Case failure
        Case vbObjectError + 304, vbObjectError + 306, 70, 3049, 3343, 3734
            BackupCopyBusy = True
        Case Else
            BackupCopyBusy = SwiftLockConflict(failure)
    End Select
End Function
' Settings & Help > Compact data file. An Access file keeps the space of deleted and
' replaced records until it is compacted. Compacting needs the data file to itself.
Public Function UXCompactDataFile() As Boolean
    Dim names As String, result As String
    On Error GoTo Failed
    If Not SwiftDataIsLinked() Then MsgBox "This workspace has no separate data file to compact.", vbInformation, "Compact data file": Exit Function
    If PresenceOthersNow(names) > 0 Then MsgBox CompactRefusal(names), vbExclamation, "Compact data file": Exit Function
    If MsgBox("Swift NAVFIT makes a verified backup beside the data file, then compacts the data file. Your other Swift NAVFIT windows close first, and nobody can open Swift NAVFIT until it finishes." & vbCrLf & vbCrLf & "Compact the data file now?", vbOKCancel + vbQuestion, "Compact data file") <> vbOK Then Exit Function
    DoCmd.Hourglass True
    result = CompactDataFile()
    DoCmd.Hourglass False
    If CurrentProject.AllForms("frmUXSettings").IsLoaded Then UXSettingsActivate
    MsgBox result, vbInformation, "Compact data file"
    UXCompactDataFile = True: Exit Function
Failed:
    DoCmd.Hourglass False
    MsgBox SwiftErrorText(), vbExclamation, "Compact data file"
End Function

Private Function CompactRefusal(ByVal names As String) As String
    CompactRefusal = "Ask everyone else to close Swift NAVFIT, then try again. Open now: " & names & ". Nothing was changed."
End Function

' Backs up the data file, sets it aside, compacts it back into its place and checks
' every table has the same number of rows; otherwise puts the original back. Returns
' what was done.
Public Function CompactDataFile() As String
    Dim path As String, folder As String, stamp As String, backup As String, old As String
    Dim names As String, before As Double, after As Double, failure As Long, message As String, released As Boolean
    Dim fso As Object, operation As Long, keeper As Long, moved As Boolean
    On Error GoTo Failed
    If Not SwiftDataIsLinked() Then Err.Raise vbObjectError + 4840, "modBackupUI.CompactDataFile", "This workspace has no separate data file to compact."
    If PresenceOthersNow(names) > 0 Then Err.Raise vbObjectError + 4841, "modBackupUI.CompactDataFile", CompactRefusal(names)
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = SwiftDataPath(): folder = fso.GetParentFolderName(path)
    stamp = Format$(Now(), "yyyymmdd_hhnnss")
    backup = folder & "\Swift_NAVFIT_backup_before_compact_" & stamp & ".accdb"
    old = Left$(path, Len(path) - 6) & " before compact " & stamp & ".accdb"
    WriteWorkspaceBackup backup
    operation = BeginOperation("Compact data file", path)
    before = fso.GetFile(path).Size
    'This session lets go of the data file: its hold, and the timer that keeps its name
    'on the list of who is connected.
    If CurrentProject.AllForms("frmSessionKeeper").IsLoaded Then keeper = Forms!frmSessionKeeper.TimerInterval: Forms!frmSessionKeeper.TimerInterval = 0
    Call SwiftDataRelease: released = True
    CompactCloseWindows
    'Set the data file aside first, so no one can open and change it while it is compacted;
    'renaming fails while anyone has it open.
    Name path As old: moved = True
    DBEngine.CompactDatabase old, path
    CompactCheckRows old, path
    after = fso.GetFile(path).Size
    On Error Resume Next
    Kill old
    On Error GoTo Failed
    Call SwiftDataHold: released = False
    If keeper > 0 Then Forms!frmSessionKeeper.TimerInterval = keeper
    FinishOperation operation, "Completed", "Compacted from " & Format$(before / 1048576#, "0.0") & " MB to " & Format$(after / 1048576#, "0.0") & " MB. Backup: " & backup
    CompactDataFile = "The data file was compacted from " & Format$(before / 1048576#, "0.0") & " MB to " & Format$(after / 1048576#, "0.0") & " MB. A verified backup was made first: " & backup
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    'Put the data file back as it was: the compacted copy goes, the original returns.
    If moved Then
        If Len(Dir$(old)) > 0 Then
            If Len(Dir$(path)) > 0 Then Kill path
            Name old As path
        End If
    End If
    If released Then SwiftDataHold
    If keeper > 0 Then Forms!frmSessionKeeper.TimerInterval = keeper
    If operation > 0 Then FinishOperation operation, "Failed", message
    On Error GoTo 0
    Select Case failure
        Case 70, 75, 3006, 3008, 3009, 3045, 3050, 3051, 3260, 3356, 3704, 3734
            message = "Someone opened Swift NAVFIT while the data file was being compacted, or it is still open elsewhere. Ask everyone else to close Swift NAVFIT, then try again. Nothing was changed."
    End Select
    Err.Raise failure, "modBackupUI.CompactDataFile", message
End Function

' Open windows keep the data file open through their lists. The backup has already
' closed the editor; every other window but Settings & Help closes too.
Private Sub CompactCloseWindows()
    Dim i As Long, name As String
    For i = Forms.Count - 1 To 0 Step -1
        name = Forms(i).Name
        If name <> "frmSessionKeeper" And name <> "frmUXSettings" Then DoCmd.Close acForm, name, acSaveNo
    Next
End Sub

' Every table in the compacted copy has as many rows as in the data file.
Private Sub CompactCheckRows(ByVal original As String, ByVal compacted As String)
    Dim a As DAO.Database, b As DAO.Database, t As DAO.TableDef, n As Long, m As Long, failure As Long, message As String
    On Error GoTo Failed
    Set a = DBEngine.OpenDatabase(original, False, True)
    Set b = DBEngine.OpenDatabase(compacted, False, True)
    For Each t In a.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" Then
            n = a.OpenRecordset("SELECT Count(*) FROM [" & t.Name & "]", dbOpenSnapshot).Fields(0).Value
            m = b.OpenRecordset("SELECT Count(*) FROM [" & t.Name & "]", dbOpenSnapshot).Fields(0).Value
            If n <> m Then Err.Raise vbObjectError + 4842, "modBackupUI.CompactCheckRows", "The compacted copy has " & m & " rows in " & t.Name & " where the data file has " & n & ". Nothing was changed."
        End If
    Next
    b.Close: a.Close
    Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not b Is Nothing Then b.Close
    If Not a Is Nothing Then a.Close
    On Error GoTo 0
    Err.Raise failure, "modBackupUI.CompactCheckRows", message
End Sub

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
' A Swift backup keeps its .verification.json beside it. When that receipt is
' present, the file must still have the hash it was verified with. Returns a line
' for the import preview, or "" when the file has no receipt.
Public Function BackupReceiptCheck(ByVal path As String, ByVal hash As String) As String
    Dim stream As Object, text As String, at As Long, verified As String
    If Len(Dir$(path & ".verification.json")) = 0 Then Exit Function
    If FileLen(path & ".verification.json") > 500000 Then Err.Raise vbObjectError + 4831, "modBackupUI.BackupReceiptCheck", "The backup's verification file is too large to read."
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile path & ".verification.json"
    text = stream.ReadText: stream.Close
    at = InStr(text, """SHA256"":""")
    If at > 0 Then verified = Mid$(text, at + 10, 64)
    If InStr(text, """Status"":""PASS""") = 0 Or Len(verified) <> 64 Then Err.Raise vbObjectError + 4831, "modBackupUI.BackupReceiptCheck", "The backup's verification file is not a passing Swift backup receipt. Restore refused; no data was changed."
    If StrComp(verified, hash, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 4832, "modBackupUI.BackupReceiptCheck", "This backup changed after it was verified: its SHA-256 no longer matches its verification file. Restore refused; no data was changed. If the change was deliberate, move the .verification.json file away and import the file as an ordinary workspace."
    BackupReceiptCheck = "Verified backup: the file matches its verification record (SHA-256 " & Left$(verified, 12) & "...)."
End Function
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
    Forms!frmBackup!lblStatus.Caption = SwiftErrorText()
End Function
