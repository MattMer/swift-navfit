Attribute VB_Name = "modBackupUI"
Option Compare Binary
Option Explicit
Private Function PowerShellLiteral(ByVal value As String) As String
    PowerShellLiteral = "'" & Replace(value, "'", "''") & "'"
End Function
Public Function EncodeWorkerCommand(ByVal command As String) As String
    Dim dom As Object, node As Object, bytes() As Byte
    bytes = command
    Set dom = CreateObject("MSXML2.DOMDocument.6.0"): Set node = dom.createElement("command")
    node.DataType = "bin.base64": node.nodeTypedValue = bytes
    EncodeWorkerCommand = Replace(Replace(node.Text, vbCr, ""), vbLf, "")
End Function
Public Function BackupWorkerCommand(ByVal mode As String, ByVal source As String, ByVal destination As String) As String
    Dim script As String
    If mode <> "Backup" And mode <> "Restore" Then Err.Raise vbObjectError + 4800, "modBackupUI.BackupWorkerCommand", "Invalid copy mode."
    script = "$ErrorActionPreference='Stop'; try { & {" & vbCrLf & BackupWorkerScript() & vbCrLf & "} -Mode " & PowerShellLiteral(mode) & " -SourcePath " & PowerShellLiteral(source) & " -DestinationPath " & PowerShellLiteral(destination) & " -WaitForCloseSeconds 30 -MaximumSchemaVersion " & SCHEMA_VERSION & " } catch { "
    script = script & "$primary=$_; $stream=$null; try { $json=@{Status='FAIL';Reason=$primary.Exception.Message;Destination=" & PowerShellLiteral(destination) & "} | ConvertTo-Json; "
    script = script & "$stream=[IO.File]::Open(" & PowerShellLiteral(destination & ".verification.json") & ",[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None); $bytes=[Text.Encoding]::UTF8.GetBytes($json); $stream.Write($bytes,0,$bytes.Length) "
    script = script & "} catch { [Console]::Error.WriteLine('Verification record not replaced: '+$_.Exception.Message) } finally { if($stream){try{$stream.Dispose()}catch{[Console]::Error.WriteLine($_.Exception.Message)}} }; [Console]::Error.WriteLine($primary.Exception.Message); exit 1 }"
    BackupWorkerCommand = "powershell.exe -NoProfile -NonInteractive -EncodedCommand " & EncodeWorkerCommand(script)
End Function
Public Function GuidedWorkspaceCopy(ByVal mode As String) As Boolean
    Dim folder As Object, sourcePicker As Object, source As String, destination As String, shell As Object, operation As Long, extension As String
    On Error GoTo Failed
    If mode = "Backup" Then
        source = CurrentProject.FullName
    ElseIf mode = "Restore" Then
        Set sourcePicker = Application.FileDialog(3): sourcePicker.Title = "Choose a closed Swift NAVFIT backup": sourcePicker.AllowMultiSelect = False
        sourcePicker.Filters.Clear: sourcePicker.Filters.Add "Swift NAVFIT database", "*.accdb;*.accde"
        If sourcePicker.Show <> -1 Then Exit Function
        source = sourcePicker.SelectedItems(1)
    Else
        Err.Raise vbObjectError + 4800, "modBackupUI.GuidedWorkspaceCopy", "Invalid copy mode."
    End If
    Set folder = Application.FileDialog(4): folder.Title = "Choose a folder for the new verified " & LCase$(mode) & " file"
    If folder.Show <> -1 Then Exit Function
    extension = Mid$(source, InStrRev(source, "."))
    destination = folder.SelectedItems(1) & "\Swift_NAVFIT_" & LCase$(mode) & "_" & Format$(Now(), "yyyymmdd_hhnnss") & "_" & Replace(Replace(NewIdentity(), "{", ""), "}", "") & extension
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Not UXCloseEditorForBackup() Then Err.Raise vbObjectError + 4801, "modBackupUI.GuidedWorkspaceCopy", "Pending editor values could not be checkpointed and closed. Save or close the editor before retrying."
    End If
    operation = BeginOperation("Guided " & LCase$(mode), destination)
    FinishOperation operation, "Dispatched", "Access will close. Read " & destination & ".verification.json for the copy verification result. Pending edits remain recovery checkpoints."
    Set shell = CreateObject("WScript.Shell")
    shell.Run BackupWorkerCommand(mode, source, destination), 0, False
    GuidedWorkspaceCopy = True
    Application.Quit acQuitSaveNone
    Exit Function
Failed:
    If CurrentProject.AllForms("frmBackup").IsLoaded Then Forms!frmBackup!lblStatus.Caption = Err.Description
End Function
Public Function ShowCopyVerification() As Boolean
    Dim picker As Object, stream As Object, path As String
    On Error GoTo Failed
    Set picker = Application.FileDialog(3): picker.Title = "Choose the backup or restore verification file": picker.AllowMultiSelect = False
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
