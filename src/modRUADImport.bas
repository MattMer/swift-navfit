Attribute VB_Name = "modRUADImport"
Option Explicit
Private rows As Collection, sourcePath As String, sourceText As String, previewStamp As String
Private prepared As Boolean
Public Function RUADReset() As Boolean
    Set rows = Nothing: sourcePath = "": sourceText = "": previewStamp = "": prepared = False
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
    prepared = False: Set rows = Nothing: previewStamp = ""
    sourcePath = path: sourceText = RUADReadText(path)
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
        Set row = CreateObject("Scripting.Dictionary")
        row.Add "Data", data: row.Add "Action", "Skip": row.Add "MemberID", 0&: row.Add "ReplaceManual", False
        row.Add "Vacant", (Len(data("NAME")) = 0 Or InStr(1, data("SELRES TYPE"), "Vacant", vbTextCompare) > 0)
        row.Add "Detail", "Choose Create new or explicitly match an existing member, then preview."
        rows.Add row
    Next
    RUADLoad = rows.Count
End Function
Public Function RUADRow(ByVal index As Long) As Object
    If rows Is Nothing Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADRow", "Choose a RUAD CSV first."
    Set RUADRow = rows(index)
End Function
Public Function RUADCount() As Long
    If Not rows Is Nothing Then RUADCount = rows.Count
End Function
Public Sub RUADChoose(ByVal index As Long, ByVal action As String, ByVal memberID As Long, Optional ByVal replaceManual As Boolean = False)
    Dim row As Object
    Set row = RUADRow(index)
    If action <> "Skip" And action <> "Create new" And action <> "Refresh existing" Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Choose a supported import action."
    If row("Vacant") And action <> "Skip" Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Vacant billets do not create directory members."
    If action = "Refresh existing" And memberID = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADChoose", "Explicitly choose the existing member."
    row("Action") = action: row("MemberID") = IIf(action = "Refresh existing", memberID, 0): row("ReplaceManual") = replaceManual
    prepared = False
End Sub
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
        Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False
        If Not doc.LoadXML(CStr(xml)) Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADSourceValues", "Existing roster provenance cannot be read."
        For Each n In doc.SelectNodes("/RUAD/Source/Field"): d.Add CStr(n.getAttribute("key")), n.Text: Next
    End If
    Set RUADSourceValues = d
End Function
Private Function RUADSourceXML(ByVal data As Object) As String
    Dim doc As Object, node As Object, field As Object, key As Variant
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False
    doc.LoadXML RUADEncode(data, CStr(data("NAME")))
    Set node = doc.createElement("Source"): node.setAttribute "path", sourcePath: node.setAttribute "importedAt", Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    For Each key In RUADKeys()
        Set field = doc.createElement("Field"): field.setAttribute "key", CStr(key): field.Text = CStr(data(key)): node.appendChild field
    Next
    doc.documentElement.appendChild node: RUADSourceXML = doc.xml
End Function
Public Function RUADPreview() As String
    Dim row As Object, incoming As Object, values As Object, baseline As Object, db As DAO.Database, r As DAO.Recordset, ids As Object, names As Object
    Dim key As Variant, name As String, rank As String, reportName As String, details As String, keep As Boolean, selected As Long, vacant As Long, conflicts As Long
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
                reportName = IIf(Len(name) <= 27, name, "")
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
                details = details & "Matched profile: " & row("Name") & ". Display name, report name, rank/rate and SSN remain unchanged." & vbCrLf
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
            row.Add "Values", values: row("SourceXML") = RUADSourceXML(incoming): row("Detail") = details
        Else
            row("Detail") = IIf(row("Vacant"), "Vacant billet: skipped; no person created.", "Skipped by selection.")
        End If
    Next
    If selected = 0 Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADPreview", "Choose at least one member to create or refresh."
    previewStamp = RUADDirectoryStamp(): prepared = True
    RUADPreview = selected & " selected changes; " & vacant & " vacant billets skipped; " & conflicts & " changed profile fields reviewed. View every selected row's full before/incoming/after values. No absent members are archived; reports and evaluation history are untouched."
End Function
Public Function RUADApply() As Long
    Dim w As DAO.Workspace, row As Object, id As Long, op As Long, failure As Long, message As String, started As Boolean
    If Not prepared Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "Preview the current selections before Apply."
    If RUADReadText(sourcePath) <> sourceText Then prepared = False: Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The CSV changed. Choose it again."
    If RUADDirectoryStamp() <> previewStamp Then prepared = False: Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The directory changed. Preview again."
    On Error GoTo Failed
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    If RUADDirectoryStamp() <> previewStamp Then Err.Raise vbObjectError + 7070, "modRUADImport.RUADApply", "The directory changed. Preview again."
    op = BeginOperation("Import RUAD", sourcePath)
    For Each row In rows
        If row("Action") <> "Skip" Then
            id = SaveDirectoryMember(CLng(row("MemberID")), CLng(row("Version")), CStr(row("Name")), CStr(row("ReportName")), CStr(row("Rank")), CStr(row("SSN")), CStr(row("SourceKind")), row("Values"), CStr(row("SourceXML")))
            RUADApply = RUADApply + 1
        End If
    Next
    FinishOperation op, "Completed", RUADApply & " reviewed roster profiles applied atomically; reports and prior history unchanged."
    w.CommitTrans: started = False: prepared = False: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRUADImport", "RUADApply", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If started Then SwiftRollback w, swiftPrimary
    prepared = False
    On Error GoTo 0
    SwiftRaise swiftPrimary & " No partial roster batch was applied."
End Function
