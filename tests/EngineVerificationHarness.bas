Attribute VB_Name = "EngineVerificationHarness"
Option Explicit
Private checks As Long, uiTicks As Long, uiBad As String, triedClose As Boolean
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise 5, , message
 checks = checks + 1
End Sub
Private Function Rejected(ByVal db As DAO.Database, ByVal table As DAO.TableDef) As Boolean
 Dim ignored As Long
 On Error GoTo Failed
 ignored = VerifyMigrationTable(db, table, "zVerifierSource")
 Exit Function
Failed:
 Rejected = (Err.Number = vbObjectError + 304 Or Err.Number = vbObjectError + 306)
End Function
Public Function TestEngineVerification() As String
 Dim db As DAO.Database, source As DAO.Database, query As DAO.QueryDef, rows As DAO.Recordset
 Dim path As String, ddl As String, body As String, changed As String, field As Variant
 On Error GoTo Failed
 Set db = CurrentDb
 path = CurrentProject.Path & "\verifier-" & Replace(Replace(NewIdentity(), "{", ""), "}", "") & ".accdb"
 Set source = DBEngine.CreateDatabase(path, dbLangGeneral, dbVersion120)
 ddl = "CREATE TABLE zVerifierFixture (ID LONG NOT NULL,BlockNumber LONG NOT NULL,TextValue TEXT(255),LongValue LONGTEXT,NullableText TEXT(120),NullNumber LONG,Stamp DATETIME,Flag BIT,Amount CURRENCY,Floating DOUBLE,CONSTRAINT PK_Probe PRIMARY KEY(ID,BlockNumber))"
 source.Execute ddl: db.Execute ddl
 source.TableDefs.Refresh: db.TableDefs.Refresh
 For Each field In Array("TextValue", "LongValue", "NullableText")
  source.TableDefs("zVerifierFixture").Fields(field).AllowZeroLength = True
  db.TableDefs("zVerifierFixture").Fields(field).AllowZeroLength = True
 Next
 body = String$(70000, "x") & ChrW(&H3A9) & ChrW(&H6F22) & vbCrLf & "Memo tail Z "
 Set rows = source.OpenRecordset("zVerifierFixture", dbOpenDynaset)
 rows.AddNew: rows!id = 1: rows!BlockNumber = 1: rows!TextValue = "Case Sensitive "
 rows!LongValue = body: rows!NullableText = Null: rows!NullNumber = Null: rows!stamp = #10/20/2026 12:34:56 PM#: rows!Flag = True: rows!Amount = 1234.5678: rows!Floating = 1.234567890123: rows.Update
 rows.AddNew: rows!id = 1: rows!BlockNumber = 2: rows!TextValue = ChrW(&HE9): rows!LongValue = "": rows!NullableText = "": rows!NullNumber = 0: rows!stamp = Null: rows!Flag = False: rows!Amount = 0: rows!Floating = 0: rows.Update: rows.Close
 Set query = db.CreateQueryDef("zVerifierSource", "SELECT * FROM zVerifierFixture IN " & Chr$(34) & path & Chr$(34)): query.Close
 db.Execute "INSERT INTO zVerifierFixture SELECT * FROM zVerifierSource", dbFailOnError
 Check VerifyMigrationTable(db, source.TableDefs("zVerifierFixture"), "zVerifierSource") = 2, "Exact composite key / long memo rejected"
 Set rows = db.OpenRecordset("SELECT * FROM zVerifierFixture WHERE BlockNumber=1", dbOpenDynaset)
 rows.Edit: rows!LongValue = Left$(body, Len(body) - 2) & "z ": rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed case change after 70,000 memo characters"
 rows.Edit: rows!LongValue = Left$(body, Len(body) - 1): rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed long-memo trailing space"
 rows.Edit: rows!LongValue = body: rows!TextValue = "case Sensitive ": rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed short text case change"
 rows.Edit: rows!TextValue = "Case Sensitive": rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed short text trailing space"
 rows.Edit: rows!TextValue = "Case Sensitive ": rows!NullableText = "": rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed Null versus empty string"
 rows.Edit: rows!NullableText = Null: rows!NullNumber = 0: rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed Null versus zero"
 rows.Edit: rows!NullNumber = Null: rows!stamp = DateAdd("s", 1, rows!stamp): rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed date/time mismatch"
 rows.Edit: rows!stamp = #10/20/2026 12:34:56 PM#: rows!Amount = 1234.5677: rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed exact currency mismatch"
 rows.Edit: rows!Amount = 1234.5678: rows!Floating = 1.234567890124: rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed floating value mismatch"
 rows.Edit: rows!Floating = 1.234567890123: rows!Flag = False: rows.Update
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed boolean mismatch"
 rows.Edit: rows!Flag = True: rows.Update: rows.Close
 db.Execute "UPDATE zVerifierFixture SET TextValue='e' WHERE BlockNumber=2", dbFailOnError
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed accent/Unicode mismatch"
 db.Execute "UPDATE zVerifierFixture SET TextValue='" & ChrW(&HE9) & "' WHERE BlockNumber=2", dbFailOnError
 db.Execute "UPDATE zVerifierFixture SET ID=2 WHERE BlockNumber=2", dbFailOnError
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed equal-count missing and extra composite key"
 db.Execute "UPDATE zVerifierFixture SET ID=1 WHERE BlockNumber=2", dbFailOnError
 db.Execute "DELETE FROM zVerifierFixture WHERE BlockNumber=2", dbFailOnError
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed missing row count"
 db.Execute "INSERT INTO zVerifierFixture SELECT * FROM zVerifierSource WHERE BlockNumber=2", dbFailOnError
 Check VerifyMigrationTable(db, source.TableDefs("zVerifierFixture"), "zVerifierSource") = 2, "Restored exact fixture rejected"
 db.Execute "INSERT INTO zVerifierFixture (ID,BlockNumber) VALUES (2,1)", dbFailOnError
 Check Rejected(db, source.TableDefs("zVerifierFixture")), "Missed extra row"
 TestEngineVerification = "PASS: " & checks & " native engine-verification adversarial checks": GoTo Cleanup
Failed:
 TestEngineVerification = "FAIL: " & Err.Number & " " & Err.Description
Cleanup:
 On Error Resume Next
 If Not rows Is Nothing Then rows.Close
 If Not source Is Nothing Then source.Close
 db.QueryDefs.Delete "zVerifierSource"
 db.Execute "DROP TABLE zVerifierFixture"
End Function
Public Function EngineImportUIPulse() As Boolean
 On Error GoTo Failed
 If Not UXWorkspaceImportBusy() Then Exit Function
 uiTicks = uiTicks + 1
 With Forms!frmUXWorkspaceImport
  If Not .Modal Or !cmdApply.Enabled Or !cmdClose.Enabled Then uiBad = "Importer was not action-locked and modal"
 End With
 If UXApplySwiftWorkspace() Then uiBad = "Nested Apply was allowed"
 If UXPreviewSwiftWorkspace(CurrentProject.FullName) Then uiBad = "Nested preview was allowed"
 If UXChooseSwiftWorkspace() Then uiBad = "Nested source chooser was allowed"
 If Not triedClose Then
  triedClose = True
  On Error Resume Next
  DoCmd.Close acForm, "frmUXWorkspaceImport", acSaveNo
  Err.Clear
  On Error GoTo Failed
  If Not CurrentProject.AllForms("frmUXWorkspaceImport").IsLoaded Then uiBad = "Busy form closed"
 End If
 EngineImportUIPulse = True: Exit Function
Failed:
 uiBad = Err.Description
End Function
Public Function EngineImportUIResult() As String
 If uiBad <> "" Then EngineImportUIResult = "FAIL: " & uiBad Else EngineImportUIResult = "PASS: " & uiTicks & " responsive event turns; nested import, preview, chooser and close blocked"
 If uiTicks = 0 Then EngineImportUIResult = "FAIL: No UI heartbeat during import"
End Function
Public Function PrepareEngineImportUI() As Boolean
 Forms!frmUXWorkspaceImport.OnTimer = "=EngineImportUIPulse()"
 Forms!frmUXWorkspaceImport.TimerInterval = 50
 PrepareEngineImportUI = True
End Function
Public Function PrepareEngineFailureUI() As Boolean
 Forms!frmUXWorkspaceImport.OnTimer = "=EngineImportFault()"
 Forms!frmUXWorkspaceImport.TimerInterval = 50
 PrepareEngineFailureUI = True
End Function
Public Function EngineImportFault() As Boolean
 Dim rows As DAO.Recordset
 On Error GoTo Finished
 If Not UXWorkspaceImportBusy() Then Exit Function
 Set rows = CurrentDb.OpenRecordset("SELECT TOP 1 ReportID FROM Report", dbOpenSnapshot)
 If rows.EOF Then rows.Close: Exit Function
 CurrentDb.Execute "UPDATE Report SET Title='Injected verification mismatch' WHERE ReportID=" & rows!ReportID, dbFailOnError
 rows.Close
 Forms!frmUXWorkspaceImport.TimerInterval = 0
 EngineImportFault = True
Finished:
End Function
