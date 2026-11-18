Attribute VB_Name = "modRSHistoryExport"
Option Compare Database
Option Explicit
' A reporting senior's RSCA history as one workbook sheet, the record a CO takes
' to the next command. For each paygrade: the BOL baseline, every report counted
' through a completed summary group (name, rank/rate, ITA), reports in groups not
' yet completed (listed, not counted), and the RSCA calculation at the bottom.
Private Const SHEET As String = "RSCA history"

Private Function Row(ByVal grade As Variant, ByVal entry As String, ByVal asOf As Variant, ByVal group As Variant, ByVal name As Variant, ByVal rank As Variant, ByVal ita As Variant, ByVal traits As Variant, ByVal traitSum As Variant, ByVal rsca As Variant, ByVal note As Variant) As Variant
    Row = Array(grade, entry, asOf, group, name, rank, ita, traits, traitSum, rsca, note)
End Function
Private Function Two(ByVal value As Double) As Double
    Two = CDbl(Format$(value, "0.00"))
End Function
Private Function Average(ByVal total As Double, ByVal count As Long) As Variant
    If count > 0 Then Average = Two(total / count) Else Average = Null
End Function

' The rows of the export, in sheet order. Each row is an 11-value array:
' Paygrade, Entry, Date, Summary group, Name, Rank or rate, ITA, Traits, Trait sum, RSCA, Note.
Public Function RSHistoryExportRows(ByVal seniorID As Long) As Collection
    Dim rows As Collection, db As DAO.Database, r As DAO.Recordset, grade As Variant, title As String
    Dim latest As Long, id As Long, parent As Long, chain As Collection, i As Long, groups As Object, key As Variant
    Dim rootCount As Long, rootSum As Double, estimated As Boolean, rootDate As Variant, rootRSCA As Variant, note As String
    Dim countedSum As Double, countedCount As Long, openSum As Double, openCount As Long, openGroups As Collection, groupTitle As String
    Dim traitSum As Double, traitCount As Long, total As Double, count As Long, recordedCount As Long, recordedSum As Variant
    Set rows = New Collection: Set db = CurrentDb
    If seniorID = 0 Or DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) = 0 Then Err.Raise vbObjectError + 4870, "modRSHistoryExport.RSHistoryExportRows", "Save the reporting senior record first."
    title = Nz(DLookup("Title", "ReportingSeniorRecord", "SeniorID=" & seniorID), "")
    rows.Add Row(Null, "Reporting senior", Date, Null, title, Null, Null, Null, Null, Null, "Swift NAVFIT local record of completed summary groups; not BOL verification.")
    For Each grade In RSTableGrades()
        latest = RSLatestBaseline(seniorID, CStr(grade))
        Set openGroups = New Collection
        Set r = db.OpenRecordset("SELECT g.GroupID FROM SummaryGroup AS g INNER JOIN GroupRSCABaseline AS b ON g.GroupID=b.GroupID WHERE b.SeniorID=" & seniorID & " AND g.Paygrade=" & QuoteSQL(CStr(grade)) & " AND Nz(g.IsArchived,False)=False ORDER BY g.Title,g.GroupID", dbOpenSnapshot)
        Do While Not r.EOF
            If RSPosting(r!GroupID) = 0 Then openGroups.Add CLng(r!GroupID)
            r.MoveNext
        Loop
        r.Close
        If latest > 0 Or openGroups.Count > 0 Then
            rows.Add Row(Null, "", Null, Null, Null, Null, Null, Null, Null, Null, Null)
            'The current history runs from the latest entry back to the entered BOL baseline.
            Set chain = New Collection: id = latest
            Do While id > 0
                If chain.Count = 0 Then chain.Add id Else chain.Add id, , 1
                parent = Nz(DLookup("ParentBaselineID", "RSCABaseline", "BaselineID=" & id), 0)
                If parent = 0 Or parent >= id Then Exit Do
                id = parent
            Loop
            rootCount = 0: rootSum = 0: estimated = False: countedSum = 0: countedCount = 0: openSum = 0: openCount = 0
            If chain.Count = 0 Then
                rows.Add Row(grade, "No BOL baseline", Null, Null, Null, Null, Null, Null, Null, Null, "No baseline is entered for this paygrade.")
            Else
                Set r = db.OpenRecordset("SELECT AsOfDate,TraitCount,TraitSum,CurrentRSCA,SourceNote FROM RSCABaseline WHERE BaselineID=" & chain(1), dbOpenSnapshot)
                rootCount = r!TraitCount: rootDate = r!AsOfDate.Value: note = Left$(Nz(r!SourceNote, ""), 200)
                If rootCount > 0 Then
                    If IsNull(r!TraitSum) Then rootSum = CDbl(Nz(r!CurrentRSCA, 0)) * rootCount: estimated = True Else rootSum = r!TraitSum
                End If
                rootRSCA = Average(rootSum, rootCount)
                If IsNull(rootRSCA) Then rootRSCA = r!CurrentRSCA.Value
                r.Close
                If estimated Then note = note & IIf(Len(note) > 0, "; ", "") & "trait sum estimated from the rounded RSCA"
                rows.Add Row(grade, "BOL baseline", rootDate, Null, Null, Null, Null, rootCount, Two(rootSum), rootRSCA, note)
            End If
            'Each group posted on the chain; a correction replaces the group's earlier posting.
            Set groups = CreateObject("Scripting.Dictionary")
            For i = 2 To chain.Count
                Set r = db.OpenRecordset("SELECT TOP 1 CompletionID,GroupID FROM GroupCompletion WHERE PostedBaselineID=" & chain(i) & " ORDER BY CompletionID", dbOpenSnapshot)
                If Not r.EOF Then groups(CStr(r!GroupID)) = CLng(r!CompletionID)
                r.Close
            Next
            For Each key In groups.Keys
                groupTitle = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & key), "Group " & key)
                Set r = db.OpenRecordset("SELECT c.TraitSum,c.TraitCount,x.OfficialName,x.MemberName,x.RankRate,x.PeriodEnd FROM GroupCompletionReport AS c LEFT JOIN Report AS x ON c.ReportID=x.ReportID WHERE c.CompletionID=" & groups(key) & " ORDER BY x.MemberName,c.ReportID", dbOpenSnapshot)
                Do While Not r.EOF
                    traitSum = Nz(r!traitSum, 0): traitCount = Nz(r!traitCount, 0)
                    countedSum = countedSum + traitSum: countedCount = countedCount + traitCount
                    rows.Add Row(grade, "Counted report", r!PeriodEnd.Value, groupTitle, UXIdentity(r!OfficialName, r!MemberName), r!RankRate.Value, Average(traitSum, traitCount), traitCount, Two(traitSum), Null, IIf(traitCount = 0, "NOB: no graded traits", Null))
                    r.MoveNext
                Loop
                r.Close
            Next
            For Each key In openGroups
                groupTitle = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & key), "Group " & key)
                Set r = RSGroupRows(CLng(key))
                Do While Not r.EOF
                    If Not Nz(r!IsArchived, False) Then
                        GradedTraits r, traitSum, traitCount
                        openSum = openSum + traitSum: openCount = openCount + traitCount
                        rows.Add Row(grade, "Not counted: group not completed", r!PeriodEnd.Value, groupTitle, UXIdentity(r!OfficialName, r!MemberName), r!RankRate.Value, Average(traitSum, traitCount), traitCount, Two(traitSum), Null, IIf(traitCount > 0, Null, IIf(Nz(r!ObservationStatus, "") = "NOB", "NOB: no graded traits", "No graded traits saved")))
                    End If
                    r.MoveNext
                Loop
                r.Close
            Next
            total = rootSum + countedSum: count = rootCount + countedCount
            note = "(" & Format$(rootSum, "0.00") & " baseline + " & Format$(countedSum, "0.00") & " counted reports) / (" & rootCount & " + " & countedCount & " traits)"
            If chain.Count > 0 Then
                recordedCount = DLookup("TraitCount", "RSCABaseline", "BaselineID=" & latest)
                recordedSum = DLookup("TraitSum", "RSCABaseline", "BaselineID=" & latest)
                If recordedCount <> count Or (Not IsNull(recordedSum) And Abs(Nz(recordedSum, 0) - total) > 0.005) Then note = note & "; differs from the recorded history (" & recordedCount & " traits): review RSCA history in Swift"
            End If
            rows.Add Row(grade, "RSCA calculation", Date, Null, Null, Null, Null, count, Two(total), Average(total, count), note)
            If openCount > 0 Then rows.Add Row(grade, "RSCA if open groups complete", Null, Null, Null, Null, Null, count + openCount, Two(total + openSum), Average(total + openSum, count + openCount), "Projection: the counted totals plus the open groups' saved traits.")
        End If
    Next
    Set RSHistoryExportRows = rows
End Function
Private Sub GradedTraits(ByVal r As DAO.Recordset, ByRef total As Double, ByRef count As Long)
    Dim key As Variant
    total = 0: count = 0
    If Nz(r!ObservationStatus, "") = "NOB" Then Exit Sub
    For Each key In TraitKeys()
        Select Case Nz(r.Fields(CStr(key)).Value, "")
            Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(r.Fields(CStr(key)).Value): count = count + 1
        End Select
    Next
End Sub
' Writes a new .xlsx workbook; never replaces an existing file. Returns the rows written.
' Access cannot write rows into a linked workbook, so the rows go into a temporary
' database and Access exports them as one sheet named RSCA history.
Public Function RSExportHistory(ByVal seniorID As Long, ByVal path As String) As Long
    Dim rows As Collection, v As Variant, db As DAO.Database, r As DAO.Recordset, q As DAO.QueryDef, i As Long, fs As Object
    Dim folder As String, staging As String, failure As Long, message As String, queryMade As Boolean, n As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    If LCase$(Right$(path, 5)) <> ".xlsx" Then Err.Raise vbObjectError + 4871, "modRSHistoryExport.RSExportHistory", "The export must be an .xlsx workbook."
    If fs.FileExists(path) Or fs.FolderExists(path) Then Err.Raise vbObjectError + 4872, "modRSHistoryExport.RSExportHistory", "A file with that name already exists. It was kept."
    If QueryExists(SHEET) Then Err.Raise vbObjectError + 4873, "modRSHistoryExport.RSExportHistory", "Another RSCA history export is in progress."
    Set rows = RSHistoryExportRows(seniorID)
    On Error GoTo Failed
    folder = fs.BuildPath(fs.GetSpecialFolder(2), "SwiftNAVFIT-rsca-" & Mid$(fs.GetTempName(), 4, 5))
    fs.CreateFolder folder: staging = fs.BuildPath(folder, "rows.accdb")
    Set db = DBEngine.CreateDatabase(staging, dbLangGeneral, dbVersion120)
    db.Execute "CREATE TABLE Rows (RowNo LONG PRIMARY KEY, Paygrade TEXT(5), Entry TEXT(60), [Date] DATETIME, [Summary group] TEXT(120), [Name] TEXT(120), [Rank or rate] TEXT(20), ITA DOUBLE, Traits LONG, [Trait sum] DOUBLE, RSCA DOUBLE, [Note] TEXT(255))", dbFailOnError
    Set r = db.OpenRecordset("Rows", dbOpenDynaset)
    For Each v In rows
        n = n + 1: r.AddNew: r!RowNo = n
        For i = 0 To 10
            If VarType(v(i)) = vbString Then
                If Len(v(i)) > 0 Then r.Fields(i + 1).Value = Left$(v(i), r.Fields(i + 1).Size)
            ElseIf Not IsNull(v(i)) Then
                r.Fields(i + 1).Value = v(i)
            End If
        Next
        r.Update
    Next
    r.Close: Set r = Nothing: db.Close: Set db = Nothing
    Set q = CurrentDb.CreateQueryDef(SHEET, "SELECT Paygrade,Entry,[Date],[Summary group],[Name],[Rank or rate],ITA,Traits,[Trait sum],RSCA,[Note] FROM Rows IN '" & Replace(staging, "'", "''") & "' ORDER BY RowNo")
    q.Close: Set q = Nothing: queryMade = True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, SHEET, path, True
    CurrentDb.QueryDefs.Delete SHEET: queryMade = False
    fs.DeleteFolder folder, True
    If Not fs.FileExists(path) Then Err.Raise vbObjectError + 4874, "modRSHistoryExport.RSExportHistory", "Access did not produce the workbook."
    RSExportHistory = rows.Count
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not db Is Nothing Then db.Close
    If queryMade Then CurrentDb.QueryDefs.Delete SHEET
    If Len(folder) > 0 Then fs.DeleteFolder folder, True
    If fs.FileExists(path) Then fs.DeleteFile path
    On Error GoTo 0
    Err.Raise failure, "modRSHistoryExport.RSExportHistory", "The RSCA history workbook could not be written: " & message
End Function
Private Function QueryExists(ByVal name As String) As Boolean
    Dim q As DAO.QueryDef
    For Each q In CurrentDb.QueryDefs
        If q.Name = name Then QueryExists = True: Exit Function
    Next
End Function

Public Function RSExportHistoryUI() As Boolean
    Dim f As Access.Form, sid As Long, dialog As Object, folder As String, path As String, n As Long
    Set f = Forms!frmUXSeniorEditor
    On Error GoTo Failed
    sid = Nz(f!txtSeniorID, 0)
    If sid = 0 Then f!lblStatus.Caption = "Save the record details before exporting the RSCA history.": Exit Function
    Set dialog = Application.FileDialog(4)
    dialog.Title = "Choose the folder for the RSCA history workbook"
    If dialog.Show <> -1 Then Exit Function
    folder = dialog.SelectedItems(1)
    path = folder & IIf(Right$(folder, 1) = "\", "", "\") & "RSCA history " & PDFOutputSafePart(Nz(DLookup("Title", "ReportingSeniorRecord", "SeniorID=" & sid), "senior")) & " " & Format$(Now, "yyyy-mm-dd hhnn") & ".xlsx"
    n = RSExportHistory(sid, path)
    f!lblStatus.Caption = "Saved " & path & ". It lists names and ITAs; store it as personnel information."
    RSExportHistoryUI = True
    Exit Function
Failed:
    f!lblStatus.Caption = "Export failed: " & Err.Description
End Function
