Attribute VB_Name = "ISICHarness"
Option Explicit
'ISICs for summary letters, the reporting senior's e-mail, and the letter's lower section.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 9990, , message
    checks = checks + 1
End Sub
Private Function Refused(ByVal title As String, ByVal uic As String, ByVal address As String) As Boolean
    On Error GoTo Failed
    Call ISICSave(0, 0, title, uic, address)
    Exit Function
Failed:
    Refused = True
End Function
Private Function EmailRefused(ByVal email As String) As Boolean
    Dim values As Object
    On Error GoTo Failed
    Set values = RSSeniorValues(0): values("ReportingSeniorName") = "BAD, E": values("ReportingSeniorEmail") = email
    Call RSSaveSenior(0, "Bad e-mail senior", values)
    Exit Function
Failed:
    EmailRefused = True
End Function
Public Function TestISIC() As String
    Dim isic As Long, other As Long, g As Long, senior As Long, baseline As Long, values As Object, details As Object, common As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC ISIC", "Tester"

    ' ISICs: name required, UIC letters and digits, four address lines at most, names unique.
    Check Refused("", "", "") And Refused("SYN ISIC", "12-45", "") And Refused("SYN ISIC", "", "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"), "Invalid ISIC accepted"
    isic = ISICSave(0, 0, "Synthetic Readiness Command", "n12345", "COMMANDER" & vbCrLf & "SYNTHETIC READINESS COMMAND" & vbCrLf & "100 EXAMPLE AVE" & vbCrLf & "EXAMPLE VA 00000")
    Check DLookup("UIC", "ISIC", "ISICID=" & isic) = "N12345", "UIC not stored in capitals"
    Check Refused("Synthetic Readiness Command", "", ""), "Duplicate ISIC name accepted"
    other = ISICSave(0, 0, "Second Synthetic ISIC", "", "")
    Call ISICSave(isic, 1, "Synthetic Readiness Command", "N12345", "COMMANDER" & vbCrLf & "SYNTHETIC READINESS COMMAND" & vbCrLf & "100 EXAMPLE AVE" & vbCrLf & "EXAMPLE VA 00000")
    Check DLookup("VersionCounter", "ISIC", "ISICID=" & isic) = 2, "ISIC edit not versioned"

    ' The Resources screen saves and lists a new ISIC; a failed save leaves no row.
    Dim rule As String, before As Long
    DoCmd.OpenForm "frmUXISICs"
    Forms!frmUXISICs!txtISICName = "Screen Synthetic ISIC": Forms!frmUXISICs!txtISICUIC = "12345": Forms!frmUXISICs!txtISICAddress = "1 SCREEN WAY"
    Check UXISICSave(), "Save ISIC failed: " & Forms!frmUXISICs!lblStatus.Caption
    Check DCount("*", "ISIC", "Title='Screen Synthetic ISIC'") = 1 And Forms!frmUXISICs!lstISICs.ListCount = DCount("*", "ISIC", "IsArchived=False"), "Saved ISIC not listed: rows " & DCount("*", "ISIC", "IsArchived=False") & ", listed " & Forms!frmUXISICs!lstISICs.ListCount
    Check UXISICNew(), "New ISIC entry did not reset"
    Forms!frmUXISICs!txtISICName = "Rolled Back ISIC"
    rule = Nz(CurrentDb.TableDefs("OperationJournal").ValidationRule, "")
    CurrentDb.TableDefs("OperationJournal").ValidationRule = "[OperationKind]<>'Save ISIC'"
    before = DCount("*", "ISIC")
    Check Not UXISICSave(), "Save with a failing operation record reported success"
    CurrentDb.TableDefs("OperationJournal").ValidationRule = rule
    Check DCount("*", "ISIC") = before And Forms!frmUXISICs!lstISICs.ListCount = DCount("*", "ISIC", "IsArchived=False"), "Failed ISIC save left a partial row"
    Check Left$(Forms!frmUXISICs!lblStatus.Caption, 11) = "Not saved: ", "Failure not reported: " & Forms!frmUXISICs!lblStatus.Caption
    Check UXISICSave() And DCount("*", "ISIC", "Title='Rolled Back ISIC'") = 1, "Retry after failure did not save"
    DoCmd.Close acForm, "frmUXISICs", acSaveNo

    ' A group remembers its ISIC; a deleted ISIC can no longer be chosen.
    g = CreateSummaryGroup("ISIC group", "EVAL", "E5")
    SetGroupISIC g, isic
    Check GroupISIC(g) = isic, "Group did not keep its ISIC"
    ISICDelete other
    On Error Resume Next
    SetGroupISIC g, other
    Check Err.Number <> 0 And GroupISIC(g) = isic, "A deleted ISIC was chosen"
    Err.Clear: On Error GoTo Failed

    ' The reporting senior's e-mail is kept on the record, not on reports.
    Check EmailRefused("not-an-address") And EmailRefused("a b@example.com"), "Invalid e-mail accepted"
    Set values = RSSeniorValues(0)
    values("ReportingSeniorName") = "SYNTHETIC, A": values("ReportingSeniorPhone") = "5550001234": values("ReportingSeniorEmail") = " senior.synthetic@example.com "
    values("ReportingSeniorAddress1") = "1 SYNTHETIC WAY": values("ReportingSeniorCity") = "EXAMPLE": values("ReportingSeniorState") = "VA": values("ReportingSeniorZip") = "000001234"
    senior = RSSaveSenior(0, "Synthetic senior", values)
    Check RSSeniorValues(senior)("ReportingSeniorEmail") = "senior.synthetic@example.com", "E-mail not saved on the senior record"
    baseline = RSAddBaseline(senior, "E5", #1/1/2026#, 4, 100, 400, "Exact fixture")
    RSSetGroupBaseline g, senior, baseline, False

    ' The letter's lower section from the group's ISIC and reporting senior.
    Set details = SummaryLetterDetails(g, isic, "No", Nothing)
    Check details("ISICAddress") = "COMMANDER" & vbCrLf & "SYNTHETIC READINESS COMMAND" & vbCrLf & "100 EXAMPLE AVE" & vbCrLf & "EXAMPLE VA 00000" And details("ISICUIC") = "N12345", "ISIC address or UIC wrong"
    Check details("SeniorAddress") = "1 SYNTHETIC WAY" & vbCrLf & "EXAMPLE, VA 00000-1234", "Senior address wrong: " & details("SeniorAddress")
    Check details("Email") = "senior.synthetic@example.com" And details("Phone") = "(555) 000-1234" And details("OnTime") = "No", "Senior contact or on-time wrong"

    ' Without a senior record, the values every report shares fill in; mixed ones stay blank.
    Set common = CreateObject("Scripting.Dictionary")
    common.Add "ReportingSeniorAddress1", "2 SHARED ST": common.Add "ReportingSeniorAddress2", "": common.Add "ReportingSeniorCity", "MIXED"
    common.Add "ReportingSeniorState", "": common.Add "ReportingSeniorZip", "": common.Add "ReportingSeniorPhone", "5550009999"
    Set details = SummaryLetterDetails(0, 0, "maybe", common)
    Check details("SeniorAddress") = "2 SHARED ST" And details("Phone") = "(555) 000-9999" And details("Email") = "" And details("ISICAddress") = "" And details("OnTime") = "", "Shared fallback wrong"
    TestISIC = "PASS: " & checks & " ISIC, senior e-mail and summary letter assertions": Exit Function
Failed:
    TestISIC = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
