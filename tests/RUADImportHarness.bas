Attribute VB_Name = "RUADImportHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7785, , message
    checks = checks + 1
End Sub
Private Sub WriteCSV(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText value: stream.SaveToFile path, 2: stream.Close
End Sub
Private Function TryPreview() As Boolean
    On Error GoTo Rejected
    Dim result As String
    result = RUADPreview(): TryPreview = True
Rejected:
End Function
Private Function TryApply() As Boolean
    On Error GoTo Rejected
    Dim result As Long
    result = RUADApply(): TryApply = True
Rejected:
End Function
Public Function TestRUADImport(ByVal sample As String) As String
    Dim i As Long, vacant As Long, selected As Long, count As Long, id As Long, initial As Long, version As Long, second As Long
    Dim row As Object, data As Object, values As Object, doc As Object, db As DAO.Database, r As DAO.Recordset, key As Variant, text As String, summary As String, path As String, failure As String, testSSN As String, stamp As String, f As Access.Form
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC RUAD TEST", "Roster tester": Set db = CurrentDb
    initial = DCount("*", "CommandMember")
    testSSN = "999888777"
    Do While DCount("*", "CommandMember", "MemberSSN=" & QuoteSQL(testSSN)) > 0: testSSN = Format$(CLng(testSSN) - 1, "000000000"): Loop
    Check RUADLoad(sample) = 66, "Sample count"
    For i = 1 To RUADCount()
        Set row = RUADRow(i)
        If row("Vacant") Then vacant = vacant + 1 Else RUADChoose i, "Create new", 0: selected = selected + 1
    Next
    Check vacant = 6 And selected = 60, "Vacancies and people"
    summary = RUADPreview(): Check InStr(summary, "60 selected") > 0, "Preview totals"
    Set row = RUADRow(1): Check InStr(row("Detail"), "Before:") > 0 And InStr(row("Detail"), "After:") > 0, "Full preview"
    Check RUADApply() = 60, "Atomic initial import"
    Check DCount("*", "CommandMember") = initial + 60, "New people count"
    Set data = row("Data"): id = CLng(DLookup("MemberID", "CommandMember", "FullName=" & QuoteSQL(CStr(data("NAME")))))
    Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot)
    Set values = RUADDecode(r!RUADData)
    For Each key In RUADKeys(): Check SameFieldValue(values(key), data(key)), "Column round trip " & key: Next
    Check IsNull(r!MemberSSN) And r!SourceKind = "RUAD", "Provisional identity"
    Check r!RankRate = data("DESG"), "Actual rate rather than billet requirement"
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.LoadXML CStr(r!RUADData)
    Check doc.SelectNodes("/RUAD/Source/Field").Length = 35, "Source baseline"
    version = r!VersionCounter: r.Close
    values("PHONE") = "Manual contact"
    Check SaveDirectoryMember(id, version, CStr(data("NAME")), "Manual report name", "IT2", testSSN, "Manual", values) = id, "Manual edit"
    text = RUADReadText(sample): path = Environ$("TEMP") & "\Swift-RUAD-test-" & NewIdentity() & ".csv"
    WriteCSV path, Replace(text, "[PII]", "Incoming contact")
    Check RUADLoad(path) = 66, "Refresh file"
    Set row = RUADRow(1): Set data = row("Data"): data("UIC") = "00001"
    RUADChoose 1, "Refresh existing", id
    summary = RUADPreview(): Check InStr(row("Detail"), "Manual contact [preserved]") > 0, "Manual conflict preview"
    Check RUADApply() = 1, "Refresh"
    Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & id, dbOpenSnapshot): Set values = RUADDecode(r!RUADData)
    Check values("PHONE") = "Manual contact", "Manual contact preserved"
    Check values("UIC") = "00001", "Identifier zeroes retained"
    Check r!MemberSSN = testSSN And r!ReportName = "Manual report name" And r!RankRate = "IT2", "Core identity preserved"
    Check DCount("*", "CommandMember") = initial + 60, "No duplicate refresh"
    version = r!VersionCounter: r.Close
    RUADChoose 1, "Refresh existing", id, True: summary = RUADPreview()
    Check RUADApply() = 1, "Explicit conflict override"
    Set values = RUADDecode(DLookup("RUADData", "CommandMember", "MemberID=" & id))
    Check values("PHONE") = "Incoming contact", "Override applied"
    RUADChoose 1, "Refresh existing", id: summary = RUADPreview()
    version = DLookup("VersionCounter", "CommandMember", "MemberID=" & id)
    stamp = RUADDirectoryStamp()
    SaveDirectoryMember id, version, CStr(data("NAME")), "Manual report name", "IT2", testSSN, "Manual", values
    Check RUADDirectoryStamp() <> stamp, "Directory stamp changed"
    Check Not TryApply(), "Stale preview rejected"
    RUADChoose 2, "Refresh existing", id: Check Not TryPreview(), "Repeated match rejected"
    RUADChoose 2, "Skip", 0: RUADChoose 1, "Create new", 0: Check Not TryPreview(), "Duplicate name rejected"
    RUADChoose 1, "Refresh existing", id
    Set row = RUADRow(2): Set data = row("Data"): data("NAME") = "Synthetic late invalid"
    RUADChoose 2, "Create new", 0
    db.TableDefs("CommandMember").Fields("FullName").ValidationRule = "<> 'Synthetic late invalid'"
    summary = RUADPreview(): version = DLookup("VersionCounter", "CommandMember", "MemberID=" & id): count = DCount("*", "CommandMember")
    Check Not TryApply(), "Late save rejected"
    Check DLookup("VersionCounter", "CommandMember", "MemberID=" & id) = version And DCount("*", "CommandMember") = count, "Whole batch rolled back"
    db.TableDefs("CommandMember").Fields("FullName").ValidationRule = ""
    RUADChoose 2, "Skip", 0: summary = RUADPreview(): WriteCSV path, text & vbCrLf
    Check Not TryApply(), "Changed source rejected"
    Check Not TryApply(), "Preview cannot apply twice"
    Check RUADCSV("A,B" & vbCrLf & """Line" & vbLf & "two"", ""a""""b""" & vbCrLf).Count = 2, "CSV multiline and escaped quotes"
    DoCmd.OpenForm "frmUXRUADImport": Set f = Forms!frmUXRUADImport
    Check RUADCount() = 0 And Not f!cmdApply.Enabled, "Fresh UI cannot reuse prior preview"
    RUADLoad sample: UXRUADImportRows
    Check f!lstRows.ListCount = 66, "All source rows visible"
    f!lstRows = 1: UXRUADImportSelect
    f!cboAction = "Refresh existing"
    Check Not UXRUADImportSelection(), "Refresh requests explicit identity"
    f!cboMatch = id
    Check UXRUADImportSelection(), "Explicit match accepted"
    Check UXRUADImportPreview() And f!cmdApply.Enabled, "Reviewed preview enables apply"
    Check InStr(f!txtDetails, "Before:") > 0 And InStr(f!txtDetails, "After:") > 0, "UI displays full values"
    f!chkReplaceManual = True: UXRUADImportSelection
    Check Not f!cmdApply.Enabled, "Changed selection invalidates UI preview"
    For i = 1 To RUADCount()
        Set row = RUADRow(i)
        If row("Vacant") Then
            f!lstRows = i: UXRUADImportSelect
            Check Not f!cboAction.Enabled And Not f!cboMatch.Enabled, "Vacancy actions disabled"
            Exit For
        End If
    Next
    DoCmd.Close acForm, "frmUXRUADImport", acSaveNo
    Kill path
    TestRUADImport = "PASS: " & checks & " RUAD mapping, identity, refresh, conflict, stale and atomic rollback assertions": Exit Function
Failed:
    failure = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    db.TableDefs("CommandMember").Fields("FullName").ValidationRule = ""
    If Len(path) > 0 Then Kill path
    TestRUADImport = failure
End Function
