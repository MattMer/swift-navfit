Attribute VB_Name = "modRUADImport"
Option Explicit
Private rows As Collection, sourcePath As String, sourceText As String, previewStamp As String
Private prepared As Boolean, previewOwner As Long, rosterComplete As Boolean
Private sourceFormat As String, rosterAsOf As Variant, rosterUIC As String, rosterPOC As String
Public Function RUADReset() As Boolean
    Set rows = Nothing: sourcePath = "": sourceText = "": previewStamp = "": prepared = False: previewOwner = 0: rosterComplete = False
    sourceFormat = "": rosterAsOf = Null: rosterUIC = "": rosterPOC = ""
    RUADReset = True
End Function
Public Function RUADReadText(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: RUADReadText = stream.ReadText: stream.Close
End Function
Public Function RUADCSV(ByVal text As String) As Collection
    Dim result As New Collection, cells As Collection, cell As String, i As Long, ch As String, quoted As Boolean, closed As Boolean
    Set cells = New Collection
    If Left$(text, 1) = ChrW(&HFEFF) Then text = Mid$(text, 2)
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If quoted Then
            If ch = """" Then
                If Mid$(text, i + 1, 1) = """" Then cell = cell & ch: i = i + 1 Else quoted = False: closed = True
            Else
                cell = cell & ch
            End If
        ElseIf ch = "," Or ch = vbCr Or ch = vbLf Then
            cells.Add Trim$(cell): cell = "": closed = False
            If ch <> "," Then
                If Not (cells.Count = 1 And cells(1) = "") Then result.Add cells
                Set cells = New Collection
                If ch = vbCr And Mid$(text, i + 1, 1) = vbLf Then i = i + 1
            End If
        ElseIf ch = """" Then
            If Len(Trim$(cell)) > 0 Or closed Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADCSV", "Unexpected CSV quote."
            cell = "": quoted = True
        Else
            If closed And ch <> " " And ch <> vbTab Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADCSV", "Unexpected text after CSV quote."
            If Not closed Then cell = cell & ch
        End If
    Next
    If quoted Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADCSV", "Unclosed CSV quote."
    If cells.Count > 0 Or Len(cell) > 0 Or closed Then
        cells.Add Trim$(cell): result.Add cells
    End If
    Set RUADCSV = result
End Function
Public Function RUADLoad(ByVal path As String) As Long
    Dim parsed As Collection, headers As Object, data As Object, row As Object, fields As Collection, key As Variant, i As Long, j As Long
    prepared = False: Set rows = Nothing: previewStamp = "": previewOwner = 0: rosterComplete = False
    sourceFormat = "NRRM": rosterAsOf = Null: rosterUIC = "": rosterPOC = ""
    sourcePath = path
    If LCase$(Right$(path, 5)) = ".xlsx" Then RUADLoad = RUADLoadMNA(path): Exit Function
    sourceText = RUADReadText(path)
    If Len(sourceText) > 5000000 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "Use a RUAD CSV smaller than 5 MB."
    Set parsed = RUADCSV(sourceText)
    If parsed.Count < 2 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "The CSV has no member rows."
    If parsed.Count > 10001 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "Use no more than 10,000 roster rows."
    Set headers = CreateObject("Scripting.Dictionary"): headers.CompareMode = vbTextCompare
    Set fields = parsed(1)
    For i = 1 To fields.Count
        If headers.Exists(fields(i)) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "Repeated CSV header: " & fields(i)
        headers.Add fields(i), i
    Next
    For Each key In RUADKeys()
        If Not headers.Exists(key) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "Missing RUAD column: " & key
    Next
    If headers.Count <> 35 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "This importer supports the verified 35-column RUAD format."
    Set rows = New Collection
    For i = 2 To parsed.Count
        Set fields = parsed(i)
        If fields.Count <> headers.Count Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "CSV row " & i & " has the wrong column count."
        Set data = CreateObject("Scripting.Dictionary")
        For Each key In RUADKeys()
            If Len(fields(headers(key))) > 4000 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoad", "CSV row " & i & ": " & key & " exceeds 4,000 characters."
            data.Add CStr(key), CStr(fields(headers(key)))
        Next
        rows.Add RUADNewRow(data, "")
    Next
    RUADLoad = rows.Count
End Function
Private Function RUADNewRow(ByVal data As Object, ByVal dodid As String) As Object
    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "Data", data: row.Add "Action", "Skip": row.Add "MemberID", 0&: row.Add "ReplaceManual", False: row.Add "Suggested", 0&: row.Add "UpdateRank", False
    row.Add "Vacant", (Len(data("NAME")) = 0 Or InStr(1, data("SELRES TYPE"), "Vacant", vbTextCompare) > 0)
    row.Add "DoDID", dodid
    row.Add "Detail", "Choose Create new or explicitly match an existing member, then preview."
    Set RUADNewRow = row
End Function
'MNA roster workbook: "Date: ..., UIC: ..." at the top, then Enlisted and Officer
'sections, each with a title row and a BIN header row. Every row with an Assign
'Type is a billet (IAP rows have no BIN). Everyone listed counts for the header UIC,
'including members cross-assigned in or out.
Private Function RUADLoadMNA(ByVal path As String) As Long
    Dim cells As Variant, kind As String, headed As Boolean, data As Object, key As Variant, n As Long
    sourceFormat = "MNA": sourceText = SourceFileHash(path)
    Set rows = New Collection
    For Each cells In RUADSheetRows(path)
        n = n + 1
        If Left$(cells(1), 5) = "Date:" And Len(rosterUIC) = 0 Then
            RUADMNAHeader cells(1)
        ElseIf Left$(cells(6), 4) = "POC:" Then
            rosterPOC = Trim$(Mid$(cells(6), 5))
            If InStr(rosterPOC, ",") > 0 Then rosterPOC = Trim$(Left$(rosterPOC, InStr(rosterPOC, ",") - 1))
        ElseIf (Left$(cells(1), 8) = "Enlisted" Or Left$(cells(1), 7) = "Officer") And Len(cells(2)) = 0 Then
            kind = Split(cells(1), " ")(0): headed = False
        ElseIf cells(1) = "BIN" Then
            If Len(kind) = 0 Or cells(8) <> "IRATE" Or cells(13) <> "DoDID" Or cells(14) <> "Name" Or cells(15) <> "PRD" Or cells(20) <> "Assign Type" Or cells(23) <> "TRUIC" Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoadMNA", "Unrecognized MNA column headers in row " & n & "."
            headed = True
        ElseIf headed And Len(cells(20)) > 0 Then
            If Len(rosterUIC) = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoadMNA", "The MNA roster has no Date and UIC line above its first section."
            Set data = CreateObject("Scripting.Dictionary")
            For Each key In RUADKeys(): data.Add CStr(key), "": Next
            data("RUIC") = rosterUIC: data("UIC") = rosterUIC: data("E/O") = kind: data("NAME") = cells(14): data("SELRES TYPE") = cells(20)
            data("DESG") = cells(8): If kind = "Enlisted" Then data("RATE/RANK") = cells(8)
            data("BIN") = cells(1): data("ABSC") = cells(2): data("UMUIC") = cells(4): data("AUIC") = cells(5): data("RBSC") = cells(6)
            data("REQ DESG/REQ RTING") = cells(7): data("RFAS") = cells(9): data("NOBC/NEC") = cells(10): data("PRD") = RUADMNAPRD(cells(15))
            data("IMS") = cells(16): data("PNOBC/PNEC") = cells(21): data("TRUIC") = cells(23)
            rows.Add RUADNewRow(data, cells(13))
            If rows.Count > 10000 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoadMNA", "Use no more than 10,000 roster rows."
        End If
    Next
    If Len(rosterUIC) = 0 Or IsNull(rosterAsOf) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoadMNA", "The MNA roster has no readable Date and UIC line."
    If rows.Count = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADLoadMNA", "The MNA roster has no billet rows."
    RUADLoadMNA = rows.Count
End Function
'Every worksheet row as cells A..W, read as text.
Private Function RUADSheetRows(ByVal path As String) As Collection
    Dim db As DAO.Database, t As DAO.TableDef, r As DAO.Recordset, sheet As String, cells() As String, i As Long, failure As Long, message As String
    Set RUADSheetRows = New Collection
    On Error GoTo Failed
    Set db = DBEngine.OpenDatabase(path, False, True, "Excel 12.0 Xml;HDR=NO;IMEX=1;")
    For Each t In db.TableDefs
        If Len(sheet) = 0 And Right$(Replace(t.Name, "'", ""), 1) = "$" Then sheet = Replace(t.Name, "'", "")
    Next
    If Len(sheet) = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADSheetRows", "The workbook has no worksheet."
    Set r = db.OpenRecordset("SELECT * FROM [" & sheet & "]", dbOpenSnapshot)
    Do While Not r.EOF
        ReDim cells(1 To 23)
        For i = 1 To 23
            If i <= r.Fields.Count Then cells(i) = Trim$(Nz(r.Fields(i - 1).Value, ""))
        Next
        RUADSheetRows.Add cells: r.MoveNext
        If RUADSheetRows.Count > 20000 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADSheetRows", "Use a workbook with no more than 20,000 rows."
    Loop
    r.Close: db.Close
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Err.Raise failure, "modRUADImport.RUADSheetRows", "The MNA workbook cannot be read: " & message
End Function
'"Date: Wed Apr 02 17:35:33 EDT 2025,  UIC: 87114"
Private Sub RUADMNAHeader(ByVal text As String)
    Dim stamp As String, parts As Variant, m As Long
    If InStr(text, ",") = 0 Or InStr(1, text, "UIC:", vbTextCompare) = 0 Then Exit Sub
    stamp = Trim$(Mid$(text, 6, InStr(text, ",") - 6))
    Do While InStr(stamp, "  ") > 0: stamp = Replace(stamp, "  ", " "): Loop
    parts = Split(stamp, " ")
    If UBound(parts) >= 3 Then
        If Len(parts(1)) = 3 Then m = InStr(1, "JanFebMarAprMayJunJulAugSepOctNovDec", parts(1), vbTextCompare)
        If m Mod 3 = 1 And IsNumeric(parts(2)) And IsNumeric(parts(UBound(parts))) Then
            If CLng(parts(UBound(parts))) >= 2000 Then rosterAsOf = DateSerial(CLng(parts(UBound(parts))), (m + 2) \ 3, CLng(parts(2)))
        End If
    End If
    rosterUIC = UCase$(Trim$(Mid$(text, InStr(1, text, "UIC:", vbTextCompare) + 4)))
End Sub
'MNA PRD is YYMM (2612 = December 2026).
Private Function RUADMNAPRD(ByVal text As String) As String
    RUADMNAPRD = text
    If text Like "####" Then
        If Val(Right$(text, 2)) >= 1 And Val(Right$(text, 2)) <= 12 Then RUADMNAPRD = "20" & Left$(text, 2) & "-" & Right$(text, 2)
    End If
End Function
'The roster's date, UIC and who pulled it (MNA); "" for an NRRM CSV.
Public Function RUADSourceSummary() As String
    If sourceFormat = "MNA" Then RUADSourceSummary = "MNA roster dated " & Format$(rosterAsOf, "yyyy-mm-dd") & " for UIC " & rosterUIC & IIf(Len(rosterPOC) > 0, ", pulled by " & rosterPOC, "") & "."
End Function
Private Function RUADSourceStamp() As String
    If sourceFormat = "MNA" Then RUADSourceStamp = SourceFileHash(sourcePath) Else RUADSourceStamp = RUADReadText(sourcePath)
End Function
Public Function RUADRow(ByVal index As Long) As Object
    If rows Is Nothing Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADRow", "Choose a RUAD CSV first."
    Set RUADRow = rows(index)
End Function
Public Function RUADCount() As Long
    If Not rows Is Nothing Then RUADCount = rows.Count
End Function
Public Sub RUADChoose(ByVal index As Long, ByVal action As String, ByVal memberID As Long, Optional ByVal replaceManual As Boolean = False, Optional ByVal updateRank As Boolean = False)
    Dim row As Object
    Set row = RUADRow(index)
    If action <> "Skip" And action <> "Create new" And action <> "Refresh existing" Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Choose a supported import action."
    If row("Vacant") And action <> "Skip" Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Vacant billets do not create directory members."
    If action = "Refresh existing" And memberID = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Explicitly choose the existing member."
    row("Action") = action: row("MemberID") = IIf(action = "Refresh existing", memberID, 0): row("ReplaceManual") = replaceManual
    row("UpdateRank") = updateRank And action = "Refresh existing"
    prepared = False
End Sub
'A complete roster lets members missing from it be flagged as departing (F02).
Public Sub RUADSetComplete(ByVal complete As Boolean)
    rosterComplete = complete: prepared = False
End Sub
Public Function RUADComplete() As Boolean
    RUADComplete = rosterComplete
End Function
'The rate/rank the roster gives a person: DESG for enlisted, RATE/RANK for officers.
Public Function RUADRosterRank(ByVal data As Object) As String
    RUADRosterRank = CStr(data("DESG"))
    If InStr(1, data("E/O"), "Officer", vbTextCompare) > 0 Then RUADRosterRank = CStr(data("RATE/RANK"))
    If Len(RUADRosterRank) > 5 Then RUADRosterRank = ""
End Function
'Suggest the existing member for each row: one active member with the row's DoD ID
'(MNA), or else whose stored roster name, or else whose display name, equals the
'row's name. Returns how many rows
'have a suggestion. A suggestion is not identity; the coordinator confirms it.
Public Function RUADSuggest() As Long
    Dim byRoster As Object, byName As Object, byDoDID As Object, r As DAO.Recordset, row As Object, name As String, values As Object, key As String
    Set byDoDID = CreateObject("Scripting.Dictionary")
    Set byRoster = CreateObject("Scripting.Dictionary"): byRoster.CompareMode = vbTextCompare
    Set byName = CreateObject("Scripting.Dictionary"): byName.CompareMode = vbTextCompare
    Set r = CurrentDb.OpenRecordset("SELECT MemberID,FullName,RUADData FROM CommandMember WHERE IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        key = Trim$(r!FullName): If byName.Exists(key) Then byName(key) = -1 Else byName.Add key, CLng(r!MemberID)
        If Not IsNull(r!RUADData) Then
            Set values = RUADDecode(r!RUADData): key = Trim$(Nz(values("NAME"), ""))
            If Len(key) > 0 Then
                If byRoster.Exists(key) Then byRoster(key) = -1 Else byRoster.Add key, CLng(r!MemberID)
            End If
            key = RUADSourceDoDID(r!RUADData)
            If Len(key) > 0 Then
                If byDoDID.Exists(key) Then byDoDID(key) = -1 Else byDoDID.Add key, CLng(r!MemberID)
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    For Each row In rows
        row("Suggested") = 0&
        If Not row("Vacant") Then
            name = Trim$(CStr(row("Data")("NAME")))
            If byDoDID.Exists(CStr(row("DoDID"))) Then
                If byDoDID(CStr(row("DoDID"))) > 0 Then row("Suggested") = byDoDID(CStr(row("DoDID")))
            ElseIf byRoster.Exists(name) Then
                If byRoster(name) > 0 Then row("Suggested") = byRoster(name)
            ElseIf byName.Exists(name) Then
                If byName(name) > 0 Then row("Suggested") = byName(name)
            End If
            If row("Suggested") > 0 Then RUADSuggest = RUADSuggest + 1
        End If
    Next
End Function
Public Function RUADDirectoryStamp() As String
    Dim db As DAO.Database, r As DAO.Recordset
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT MemberID,VersionCounter,IsArchived FROM CommandMember ORDER BY MemberID", dbOpenSnapshot)
    Do While Not r.EOF
        RUADDirectoryStamp = RUADDirectoryStamp & r!MemberID & ":" & r!VersionCounter & ":" & r!IsArchived & ";": r.MoveNext
    Loop
    r.Close
End Function
Private Function RUADSourceValues(ByVal xml As Variant) As Object
    Dim doc As Object, d As Object, n As Object
    Set d = CreateObject("Scripting.Dictionary")
    If Len(Nz(xml, "")) > 0 Then
        Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
        If Not doc.LoadXML(CStr(xml)) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADSourceValues", "Existing roster provenance cannot be read."
        For Each n In doc.SelectNodes("/RUAD/Source/Field"): d.Add CStr(n.getAttribute("key")), n.Text: Next
    End If
    Set RUADSourceValues = d
End Function
'The DoD ID recorded from an MNA roster, or "".
Private Function RUADSourceDoDID(ByVal xml As Variant) As String
    Dim doc As Object, node As Object
    If InStr(1, Nz(xml, ""), "dodid=", vbTextCompare) = 0 Then Exit Function
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
    If Not doc.LoadXML(CStr(xml)) Then Exit Function
    Set node = doc.SelectSingleNode("/RUAD/Source/@dodid")
    If Not node Is Nothing Then RUADSourceDoDID = node.Text
End Function
Private Function RUADSourceXML(ByVal data As Object, ByVal dodid As String) As String
    Dim doc As Object, node As Object, field As Object, key As Variant
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
    doc.LoadXML RUADEncode(data, CStr(data("NAME")))
    Set node = doc.createElement("Source"): node.setAttribute "path", sourcePath: node.setAttribute "importedAt", Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    If sourceFormat = "MNA" Then node.setAttribute "format", "MNA": node.setAttribute "asOf", Format$(rosterAsOf, "yyyy-mm-dd")
    If Len(dodid) > 0 Then node.setAttribute "dodid", dodid
    For Each key In RUADKeys()
        Set field = doc.createElement("Field"): field.setAttribute "key", CStr(key): field.Text = CStr(data(key)): node.appendChild field
    Next
    doc.documentElement.appendChild node: RUADSourceXML = doc.xml
End Function
Public Function RUADPreview() As String
    Dim row As Object, incoming As Object, values As Object, baseline As Object, db As DAO.Database, r As DAO.Recordset, ids As Object, names As Object
    Dim key As Variant, name As String, rank As String, reportName As String, details As String, keep As Boolean, selected As Long, vacant As Long, conflicts As Long, summary As String
    prepared = False: Set db = CurrentDb: Set ids = CreateObject("Scripting.Dictionary"): Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbTextCompare
    If rows Is Nothing Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Choose a CSV first."
    For Each row In rows
        Set incoming = row("Data")
        If row("Vacant") Then vacant = vacant + 1
        If row("Action") <> "Skip" Then
            selected = selected + 1: name = CStr(incoming("NAME"))
            If name = "" Or Len(name) > 120 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Member name must fit within 120 characters."
            details = row("Action") & " | " & name & vbCrLf
            If row("Action") = "Create new" Then
                If names.Exists(name) Or DCount("*", "CommandMember", "FullName=" & QuoteSQL(name)) > 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "An existing or repeated name requires explicit matching: " & name
                names.Add name, True: Set values = RUADDecode(Null): Set baseline = CreateObject("Scripting.Dictionary")
                rank = CStr(incoming("DESG"))
                If InStr(1, incoming("E/O"), "Officer", vbTextCompare) > 0 Then rank = CStr(incoming("RATE/RANK"))
                If Len(rank) > 5 Then rank = ""
                reportName = name
                'MNA names are LAST FIRST MIDDLE without the comma.
                If sourceFormat = "MNA" Then reportName = NavyNameFormat(Replace(name, " ", ", ", 1, 1))
                If Len(reportName) > 27 Then reportName = ""
                row("Version") = 0&: row("Name") = name: row("Rank") = rank: row("ReportName") = reportName: row("SSN") = "": row("SourceKind") = "RUAD"
                details = details & "New provisional member; SSN unentered. Report rank/rate: " & rank & "; official name: " & reportName & vbCrLf
            Else
                If ids.Exists(CStr(row("MemberID"))) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Two CSV rows cannot refresh the same member."
                ids.Add CStr(row("MemberID")), True
                Set r = db.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & CLng(row("MemberID")), dbOpenSnapshot)
                If r.EOF Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Matched member no longer exists."
                If r!IsArchived Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Choose an active member."
                Set values = RUADDecode(r!RUADData): Set baseline = RUADSourceValues(r!RUADData)
                row("Version") = CLng(r!VersionCounter): row("Name") = CStr(r!FullName): row("Rank") = Nz(r!RankRate, ""): row("ReportName") = Nz(r!ReportName, ""): row("SSN") = Nz(r!MemberSSN, ""): row("SourceKind") = Nz(r!SourceKind, "Manual")
                details = details & "Matched profile: " & row("Name") & ". Display name, report name and SSN remain unchanged." & vbCrLf
                If row("UpdateRank") And Len(RUADRosterRank(incoming)) > 0 Then
                    details = details & "Rank/rate: " & row("Rank") & " -> " & RUADRosterRank(incoming) & " (updated from this roster)" & vbCrLf: row("Rank") = RUADRosterRank(incoming)
                ElseIf RUADRosterRank(incoming) <> row("Rank") And Len(RUADRosterRank(incoming)) > 0 Then
                    details = details & "Rank/rate stays " & row("Rank") & "; this roster shows " & RUADRosterRank(incoming) & ". Tick Update the profile rank/rate to change it." & vbCrLf
                End If
                r.Close: Set r = Nothing
            End If
            For Each key In RUADKeys()
                keep = False
                If row("Action") = "Refresh existing" Then
                    If key = "NAME" Then
                        keep = True
                    ElseIf baseline.Exists(key) Then
                        keep = Not SameFieldValue(values(key), baseline(key)) And Not SameFieldValue(values(key), incoming(key))
                    Else
                        keep = Len(Nz(values(key), "")) > 0 And Not SameFieldValue(values(key), incoming(key))
                    End If
                    If keep And key <> "NAME" Then conflicts = conflicts + 1: keep = Not row("ReplaceManual")
                End If
                details = details & CStr(key) & vbCrLf & "  Before: " & Nz(values(key), "(unentered)") & vbCrLf & "  Incoming: " & incoming(key) & vbCrLf
                If Not keep Then values(key) = incoming(key)
                If key = "NAME" Then values(key) = row("Name")
                details = details & "  After: " & Nz(values(key), "(unentered)") & IIf(keep, " [preserved]", "") & vbCrLf
            Next
            If row.Exists("Values") Then row.Remove "Values"
            row.Add "Values", values: row("SourceXML") = RUADSourceXML(incoming, CStr(row("DoDID"))): row("Detail") = details
        Else
            row("Detail") = IIf(row("Vacant"), "Vacant billet: skipped; no person created.", "Skipped by selection.")
        End If
    Next
    If selected = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Choose at least one member to create or refresh."
    previewStamp = RUADDirectoryStamp(): previewOwner = ActiveContributorID(): prepared = True
    If rosterComplete Then summary = " Complete roster: active members on the last complete roster for its UICs who are missing now will be flagged as departing." Else summary = " Not marked as a complete roster, so no one is flagged as missing."
    RUADPreview = selected & " selected changes; " & vacant & " vacant billets skipped; " & conflicts & " changed profile fields reviewed. View every selected row's full before/incoming/after values. No absent members are archived; reports and evaluation history are untouched." & summary
End Function
Public Function RUADApply() As Long
    Dim w As DAO.Workspace, row As Object, id As Long, op As Long, failure As Long, message As String, started As Boolean
    If Not prepared Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "Preview the current selections before Apply."
    If RUADSourceStamp() <> sourceText Then prepared = False: Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The roster file changed. Choose it again."
    If RUADDirectoryStamp() <> previewStamp Then prepared = False: Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The directory changed. Preview again."
    If previewOwner <> ActiveContributorID() Then prepared = False: Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "Preview the roster changes as the current contributor before Apply."
    On Error GoTo Failed
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    If RUADDirectoryStamp() <> previewStamp Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The directory changed. Preview again."
    op = BeginOperation("Import RUAD", sourcePath)
    Dim present As Object, scope As String, data As Object, member As Long
    Set present = CreateObject("Scripting.Dictionary"): scope = "|"
    For Each row In rows
        Set data = row("Data"): member = 0
        If row("Action") <> "Skip" Then
            id = SaveDirectoryMember(CLng(row("MemberID")), CLng(row("Version")), CStr(row("Name")), CStr(row("ReportName")), CStr(row("Rank")), CStr(row("SSN")), CStr(row("SourceKind")), row("Values"), CStr(row("SourceXML")))
            RUADApply = RUADApply + 1: member = id
        ElseIf Not row("Vacant") Then
            'A skipped row for a known member still shows they are on the roster.
            member = CLng(row("Suggested"))
        End If
        If Len(Trim$(CStr(data("UIC")))) > 0 And InStr(scope, "|" & Trim$(CStr(data("UIC"))) & "|") = 0 Then scope = scope & Trim$(CStr(data("UIC"))) & "|"
        If member > 0 And Not present.Exists(CStr(member)) Then present.Add CStr(member), Array(Trim$(CStr(data("UIC"))), Trim$(CStr(data("PRD"))), RUADRosterRank(data))
    Next
    Call RecordRosterImport(CurrentDb, sourceFormat, Mid$(sourcePath, InStrRev(sourcePath, "\") + 1), SourceFileHash(sourcePath), scope, rosterComplete, rows.Count, present, rosterAsOf)
    FinishOperation op, "Completed", RUADApply & " reviewed roster profiles applied atomically; reports and prior history unchanged." & IIf(sourceFormat = "MNA", " " & RUADSourceSummary(), "")
    w.CommitTrans: started = False: prepared = False: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRUADImport", "RUADApply", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If started Then SwiftRollback w, swiftPrimary
    prepared = False
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
