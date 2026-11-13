Attribute VB_Name = "modWorkspaceSafety"
Option Compare Binary
Option Explicit

'Startup warnings for situations the one-file, one-person workspace model cannot
'protect: a file Windows may delete, a folder a sync service rewrites, a file that
'cannot be saved, and a second session on the same file. They warn; they never block.

Private Const JET_USER_ROSTER As String = "{947BB102-5D43-11D1-BDBF-00C04FB92675}"
Private Const ATTRIBUTE_REPARSE_POINT As Long = 1024

'Path-only checks, so tests can classify synthetic paths without creating files.
Public Function WorkspaceLocationWarning(ByVal path As String) As String
    Dim p As String, text As String
    p = LCase$(Replace(path, "/", "\"))
    If PathUnder(p, Environ$("TEMP")) Or PathUnder(p, Environ$("TMP")) Or InStr(p, "\appdata\local\temp\") > 0 Or InStr(p, "\inetcache\") > 0 Or InStr(p, "\temporary internet files\") > 0 Or InStr(p, "\content.outlook\") > 0 Then
        text = "This file is open from a temporary folder, for example from inside a zip file or an e-mail attachment. Windows can delete it, and your work with it. Exit, save the file to an approved folder and open it from there."
    End If
    If PathUnder(p, Environ$("OneDrive")) Or PathUnder(p, Environ$("OneDriveCommercial")) Or PathUnder(p, Environ$("OneDriveConsumer")) Or InStr(p, "\onedrive\") > 0 Or InStr(p, "\onedrive - ") > 0 Or InStr(p, "\dropbox\") > 0 Or InStr(p, "\google drive\") > 0 Or InStr(p, "\my drive\") > 0 Or InStr(p, "\icloud drive\") > 0 Then
        text = WarningJoin(text, SyncWarning())
    End If
    WorkspaceLocationWarning = text
End Function

Private Function SyncWarning() As String
    SyncWarning = "This file is in a folder that a sync service such as OneDrive or SharePoint uploads. Sync can damage an open Access database or create conflicting copies. Keep the workspace in a folder that is not synced."
End Function

Private Function PathUnder(ByVal path As String, ByVal folder As String) As Boolean
    folder = LCase$(Replace(Trim$(folder), "/", "\"))
    If Len(folder) = 0 Then Exit Function
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    PathUnder = (Left$(path, Len(folder)) = folder)
End Function

Private Function WarningJoin(ByVal text As String, ByVal addition As String) As String
    If Len(addition) = 0 Then WarningJoin = text: Exit Function
    If Len(text) = 0 Then WarningJoin = addition Else WarningJoin = text & vbCrLf & addition
End Function

'Cloud placeholders (OneDrive, SharePoint libraries) are reparse points even when
'the content is local; a path name alone misses libraries synced under other names.
Public Function WorkspaceFileWarning(ByVal path As String) As String
    Dim fso As Object, attributes As Long, text As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    attributes = fso.GetFile(path).attributes
    If (attributes And vbReadOnly) <> 0 Then text = ReadOnlyWarning()
    If (attributes And ATTRIBUTE_REPARSE_POINT) <> 0 Then
        If InStr(WorkspaceLocationWarning(path), SyncWarning()) = 0 Then text = WarningJoin(text, SyncWarning())
    End If
    WorkspaceFileWarning = text
End Function

Private Function ReadOnlyWarning() As String
    ReadOnlyWarning = "This workspace is read-only: the file or its folder cannot be written. Nothing you enter can be saved. Exit, copy the file to a folder you can write to and open it from there."
End Function

'Each open session holds one connected slot in the Access lock file. This session's
'own slot is one of them; Access reports the computer name, not the person.
Public Function OtherWorkspaceSessions(ByRef computers As String) As Long
    Dim rows As Object, total As Long, name As String, own As String, ownSeen As Boolean
    computers = ""
    'With a separate data file the workspace is that file, and it has no session list
    'to read: startup records whether another session was holding it (SwiftDataHold).
    If SwiftDataIsLinked() Then
        If SwiftDataOtherSession() Then OtherWorkspaceSessions = 1
        Exit Function
    End If
    own = UCase$(Environ$("COMPUTERNAME"))
    'An exclusive open has no shared roster (ADO error 3251) and, by definition,
    'no other session. Any other probe failure reaches the caller.
    On Error GoTo NoRoster
    Set rows = CurrentProject.Connection.OpenSchema(-1, , JET_USER_ROSTER)
    On Error GoTo 0
    Do While Not rows.EOF
        If CBool(Nz(rows.Fields("CONNECTED").Value, False)) Then
            total = total + 1
            name = UCase$(Trim$(Replace(Nz(rows.Fields("COMPUTER_NAME").Value, ""), Chr$(0), "")))
            If name = own And Not ownSeen Then
                ownSeen = True
            Else
                If name = own Then name = "this computer"
                If InStr(", " & computers & ",", ", " & name & ",") = 0 Then
                    If Len(computers) > 0 Then computers = computers & ", "
                    computers = computers & name
                End If
            End If
        End If
        rows.MoveNext
    Loop
    rows.Close
    If total > 0 Then OtherWorkspaceSessions = total - 1
    Exit Function
NoRoster:
    If Err.Number = 3251 Then Exit Function
    Err.Raise Err.Number, "modWorkspaceSafety.OtherWorkspaceSessions", Err.Description
End Function

Public Function ConcurrentSessionWarning(ByVal others As Long, ByVal computers As String) As String
    If others <= 0 Then Exit Function
    ConcurrentSessionWarning = "This workspace is already open in " & IIf(others = 1, "another session", others & " other sessions") & IIf(Len(computers) > 0, " (" & computers & ")", "") & ". Swift NAVFIT supports one person at a time: two sessions change each other's contributor and can damage the file. Close the other session first, or exit."
End Function

'All warnings for the open workspace, one per line. A failed probe is reported as
'its own warning rather than hidden.
Public Function WorkspaceSafetyNotice() As String
    Dim text As String, others As Long, computers As String, path As String
    path = SwiftDataPath()
    text = WorkspaceLocationWarning(path)
    On Error GoTo FileFailed
    text = WarningJoin(text, WorkspaceFileWarning(path))
    If Not CurrentDb.Updatable And InStr(text, ReadOnlyWarning()) = 0 Then text = WarningJoin(text, ReadOnlyWarning())
SessionCheck:
    On Error GoTo SessionFailed
    others = OtherWorkspaceSessions(computers)
    text = WarningJoin(text, ConcurrentSessionWarning(others, computers))
Finished:
    WorkspaceSafetyNotice = text
    Exit Function
FileFailed:
    text = WarningJoin(text, "Swift NAVFIT could not check this file's attributes (error " & Err.Number & ": " & Err.Description & ").")
    Resume SessionCheck
SessionFailed:
    text = WarningJoin(text, "Swift NAVFIT could not check whether this workspace is open elsewhere (error " & Err.Number & ": " & Err.Description & ").")
    Resume Finished
End Function

'Paint the warning panel on a startup form. The panel is hidden when nothing applies.
Public Function UXWorkspaceSafetyPaint(ByVal formName As String) As Boolean
    Dim c As Access.Control, text As String
    Set c = Forms(formName).Controls("uxWorkspaceWarning")
    text = WorkspaceSafetyNotice()
    c.Caption = text
    c.Visible = (Len(text) > 0)
    If Len(text) > 0 Then
        c.BackStyle = 1: c.BackColor = UXColor("Warning"): c.ForeColor = UXColor("WarningInk")
    End If
    UXWorkspaceSafetyPaint = True
End Function
