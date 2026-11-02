Attribute VB_Name = "NAVFITUpgradeHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8062, , message
 n = n + 1
End Sub
Private Function Rejected(ByVal path As String) As Boolean
 On Error GoTo Expected
 OpenReviewFile path, False
 CloseReviewFile
 Exit Function
Expected:
 Rejected = True
End Function
Public Function TestNAVFITUpgrade(ByVal directory As String, ByVal template As String) As String
 On Error GoTo Failed
 Dim file As Variant, path As String, hash As String, copy As String, values As Object
 Dim entry As Variant, parts As Variant, db As DAO.Database, r As DAO.Recordset, w As DAO.Workspace
 Dim id As Long, nativeID As Long, groups As Object, picks As Collection, result As Collection, key As Variant, preview As Object
 Dim baseCount As Long, failed As Boolean, tempHash As String
 ConfigureWorkspace "SYNTHETIC V33 UPGRADE", "Tester"
 For Each file In Array("modern", "legacy", "partial", "readonly")
  path = directory & "\" & file & ".accdb": hash = SourceFileHash(path)
  Check OpenReviewFile(path, False), "Open failed: " & file
  copy = ReviewPreparedCopyPath()
  If file = "modern" Then
   Check copy = "" And ReviewUpgradeMessage() = "", "Current schema was converted"
  Else
   Check copy <> "" And copy <> path, "Legacy file not isolated"
   Check InStr(ReviewUpgradeMessage(), "upgraded automatically") > 0, "Upgrade not disclosed"
   Set db = DBEngine.OpenDatabase(copy, False, True): ValidateNAVFITDatabase db
   For Each entry In NAVFITPatchColumns()
    parts = Split(entry, "|")
    Check db.TableDefs(parts(0)).Fields(parts(1)).Type = dbText, "Patch type differs"
    Check db.TableDefs(parts(0)).Fields(parts(1)).Size = CLng(parts(2)), "Patch size differs"
    Set r = db.OpenRecordset("SELECT Count(*) FROM [" & parts(0) & "] WHERE [" & parts(1) & "] IS NULL", dbOpenSnapshot)
    Check r.Fields(0) = 0, "Patch NULL not initialized": r.Close
   Next
   If file = "partial" Then
    Set r = db.OpenRecordset("SELECT RSPhone FROM Folders WHERE RSPhone='5551234567'", dbOpenSnapshot)
    Check Not r.EOF, "Existing patch value overwritten": r.Close
   End If
   db.Close: Set db = Nothing
  End If
  Set groups = NAVFITReviewGroupChoices(2000): Check groups.Count = 3, "Group preview count differs"
  For Each key In groups.Keys
   Check groups(key)("SourceHash") = hash, "Preview lost original hash"
  Next
  Check SourceFileHash(path) = hash, "Source changed while preparing"
  CloseReviewFile
  If copy <> "" Then Check Dir$(copy) = "", "Temporary copy leaked after close"
 Next
 For Each file In Array("bad-size", "missing-core", "unknown-column", "failed-init")
  Check Rejected(directory & "\" & file & ".accdb"), "Unsupported or failed schema accepted: " & file
 Next
 Check DCount("*", "Report") = 0 And DCount("*", "ImportReceipt") = 0, "Failed open wrote local data"
 path = directory & "\legacy.accdb": hash = SourceFileHash(path)
 Check OpenReviewFile(path, False), "Legacy reopen failed"
 Set groups = NAVFITReviewGroupChoices(2000): Set picks = New Collection
 For Each key In groups.Keys: picks.Add CStr(key): Next
 Set result = ImportNAVFITGroups(picks, hash, 2000)
 Check result.Count = 3 And DCount("*", "Report") = 3, "Full group import failed"
 Check DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(hash) & " AND SourcePath=" & QuoteSQL(path)) = 3, "Receipt lost original path/hash"
 For Each key In result
  Set preview = NAVFITGroupPreview(CLng(key))
  Call NAVFITExportGroup(CLng(key), directory & "\export-" & key & ".accdb", template, preview("Token"))
  Set db = DBEngine.OpenDatabase(directory & "\export-" & key & ".accdb", False, True)
  ValidateNAVFITDatabase db: Check Not NAVFITNeedsV33Upgrade(db), "Export lost Patch 2 schema": db.Close
 Next
 Call CloseReviewFile: Check SourceFileHash(path) = hash, "Import/export altered original"
 Check OpenReviewFile(path, False), "Repeated conversion failed"
 Set groups = NAVFITReviewGroupChoices(2000): failed = False
 On Error Resume Next
 Set result = ImportNAVFITGroups(picks, hash, 2000)
 failed = (Err.Number <> 0): Err.Clear
 On Error GoTo Failed
 Check failed And DCount("*", "Report") = 3, "Repeated conversion bypassed duplicates"
 CloseReviewFile
 path = directory & "\stale.accdb": Check OpenReviewFile(path, False), "Stale fixture failed"
 copy = ReviewPreparedCopyPath(): Set db = DBEngine.OpenDatabase(path)
 db.Execute "UPDATE Reports SET FullName='CHANGED SOURCE' WHERE ReportType='Eval'", dbFailOnError: db.Close
 failed = False: On Error Resume Next
 Set groups = NAVFITReviewGroupChoices(2000): failed = (Err.Number <> 0): Err.Clear
 On Error GoTo Failed
 Check failed, "Changed original accepted": CloseReviewFile
 path = directory & "\legacy.accdb": Check OpenReviewFile(path, False), "Tamper fixture failed"
 copy = ReviewPreparedCopyPath(): Set db = DBEngine.OpenDatabase(copy)
 db.Execute "UPDATE Reports SET FullName='CHANGED COPY' WHERE ReportType='Eval'", dbFailOnError: db.Close
 failed = False: On Error Resume Next
 Set groups = NAVFITReviewGroupChoices(2000): failed = (Err.Number <> 0): Err.Clear
 On Error GoTo Failed
 Check failed, "Changed prepared copy accepted": CloseReviewFile
 TestNAVFITUpgrade = "PASS: " & n & " legacy/partial/current/read-only upgrades, guards, provenance, group import and export assertions"
 Exit Function
Failed:
 TestNAVFITUpgrade = "FAIL: " & file & " " & Err.Number & " " & Err.Description
 On Error Resume Next
 If Not db Is Nothing Then db.Close
 CloseReviewFile
End Function
