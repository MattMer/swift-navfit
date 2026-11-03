Attribute VB_Name = "modISIC"
Option Explicit
'ISICs (immediate superiors in command) for summary letters: kept under Resources,
'chosen per summary group when its PDFs are created.
Private isicID As Long, isicVersion As Long

Public Function ISICSave(ByVal id As Long, ByVal expectedVersion As Long, ByVal title As String, ByVal uic As String, ByVal address As String) As Long
    Dim r As DAO.Recordset, operation As Long
    title = Trim$(title): uic = UCase$(Trim$(uic)): address = Trim$(address)
    If Len(title) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 9901, "modISIC.ISICSave", "Enter the ISIC's name, up to 120 characters."
    If Len(uic) > 6 Or uic Like "*[!A-Z0-9]*" Then Err.Raise vbObjectError + 9902, "modISIC.ISICSave", "The UIC is up to 6 letters and digits."
    If Len(address) > 400 Then Err.Raise vbObjectError + 9903, "modISIC.ISICSave", "Keep the address within 400 characters."
    If UBound(Split(Replace(address, vbCrLf, vbLf), vbLf)) > 3 Then Err.Raise vbObjectError + 9903, "modISIC.ISICSave", "The summary letter has room for four address lines."
    If DCount("*", "ISIC", "IsArchived=False AND ISICID<>" & id & " AND Title=" & QuoteSQL(title)) > 0 Then Err.Raise vbObjectError + 9904, "modISIC.ISICSave", "Another ISIC already has this name."
    If id = 0 Then
        Set r = CurrentDb.OpenRecordset("ISIC", dbOpenDynaset)
        r.AddNew: r!IsArchived = False: r!VersionCounter = 1
    Else
        Set r = CurrentDb.OpenRecordset("SELECT * FROM ISIC WHERE ISICID=" & id, dbOpenDynaset)
        If r.EOF Then r.Close: Err.Raise vbObjectError + 9905, "modISIC.ISICSave", "This ISIC no longer exists."
        If r!VersionCounter <> expectedVersion Then r.Close: Err.Raise vbObjectError + 9906, "modISIC.ISICSave", "This ISIC changed. Choose it again before saving."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!Title = title: r!UIC = IIf(Len(uic) = 0, Null, uic): r!Address = IIf(Len(address) = 0, Null, address)
    r.Update: r.Bookmark = r.LastModified: ISICSave = r!ISICID: r.Close
    operation = BeginOperation("Save ISIC", title)
    FinishOperation operation, "Completed", "ISIC saved for summary letters."
End Function
'Deleting hides the ISIC from new choices; groups that used it keep it on record.
Public Sub ISICDelete(ByVal id As Long)
    Dim title As String, operation As Long
    title = Nz(MajorScalar("SELECT Title FROM ISIC WHERE IsArchived=False AND ISICID=" & id), "")
    If Len(title) = 0 Then Err.Raise vbObjectError + 9905, "modISIC.ISICDelete", "Choose an ISIC to delete."
    operation = BeginOperation("Delete ISIC", title)
    CurrentDb.Execute "UPDATE ISIC SET IsArchived=True,VersionCounter=VersionCounter+1 WHERE ISICID=" & id, dbFailOnError
    FinishOperation operation, "Completed", "ISIC hidden from new summary letters."
End Sub
Public Sub SetGroupISIC(ByVal groupID As Long, ByVal id As Long)
    If groupID <= 0 Then Exit Sub
    If id > 0 Then
        If DCount("*", "ISIC", "IsArchived=False AND ISICID=" & id) = 0 Then Err.Raise vbObjectError + 9905, "modISIC.SetGroupISIC", "Choose an active ISIC."
    End If
    CurrentDb.Execute "UPDATE SummaryGroup SET ISICID=" & IIf(id > 0, CStr(id), "Null") & " WHERE GroupID=" & groupID, dbFailOnError
End Sub
Public Function GroupISIC(ByVal groupID As Long) As Long
    If groupID > 0 Then GroupISIC = Nz(MajorScalar("SELECT ISICID FROM SummaryGroup WHERE GroupID=" & groupID), 0)
End Function

'The summary letter's lower section for a group: ISIC address and UIC, the reporting
'senior's address, e-mail and phone, and whether the report was submitted on time.
'Senior details come from the group's reporting senior record, then from the values
'every report in the group shares.
Public Function SummaryLetterDetails(ByVal groupID As Long, ByVal chosenISIC As Long, ByVal onTime As String, ByVal common As Object) As Object
    Dim result As Object, senior As Long, values As Object, key As Variant, r As DAO.Recordset, city As String
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "ISICAddress", "": result.Add "ISICUIC", "": result.Add "SeniorAddress", "": result.Add "Email", "": result.Add "Phone", ""
    result.Add "OnTime", IIf(onTime = "Yes" Or onTime = "No", onTime, "")
    If chosenISIC > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT Title,UIC,Address FROM ISIC WHERE ISICID=" & chosenISIC, dbOpenSnapshot)
        If Not r.EOF Then
            result("ISICAddress") = Trim$(Nz(r!Address, "")): result("ISICUIC") = Nz(r!UIC, "")
            If Len(result("ISICAddress")) = 0 Then result("ISICAddress") = r!Title
        End If
        r.Close
    End If
    If groupID > 0 Then senior = Nz(MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & groupID), 0)
    If senior > 0 Then
        Set values = RSSeniorValues(senior)
    Else
        Set values = CreateObject("Scripting.Dictionary")
        For Each key In RSRecordKeys(): values.Add CStr(key), Null: Next
    End If
    For Each key In Array("ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone")
        If IsNull(values(CStr(key))) And Not common Is Nothing Then
            If common.Exists(CStr(key)) Then
                If common(CStr(key)) <> "MIXED" And common(CStr(key)) <> "" Then values(CStr(key)) = common(CStr(key))
            End If
        End If
    Next
    city = Trim$(Nz(values("ReportingSeniorCity"), ""))
    If Len(Nz(values("ReportingSeniorState"), "")) > 0 Then city = city & IIf(Len(city) > 0, ", ", "") & values("ReportingSeniorState")
    If Len(Nz(values("ReportingSeniorZip"), "")) > 0 Then city = city & IIf(Len(city) > 0, " ", "") & LetterZip(CStr(values("ReportingSeniorZip")))
    result("SeniorAddress") = LetterLines(Array(Nz(values("ReportingSeniorAddress1"), ""), Nz(values("ReportingSeniorAddress2"), ""), city))
    result("Email") = Nz(values("ReportingSeniorEmail"), "")
    result("Phone") = LetterPhone(Nz(values("ReportingSeniorPhone"), ""))
    Set SummaryLetterDetails = result
End Function
Private Function LetterLines(ByVal parts As Variant) As String
    Dim part As Variant
    For Each part In parts
        If Len(Trim$(part)) > 0 Then LetterLines = LetterLines & IIf(Len(LetterLines) > 0, vbCrLf, "") & Trim$(part)
    Next
End Function
'(xxx) xxx-xxxx for ten digits, as NAVFIT98A prints it; anything else as entered.
Private Function LetterPhone(ByVal phone As String) As String
    LetterPhone = phone
    If phone Like "##########" Then LetterPhone = "(" & Left$(phone, 3) & ") " & Mid$(phone, 4, 3) & "-" & Right$(phone, 4)
End Function
Private Function LetterZip(ByVal zip As String) As String
    LetterZip = zip
    If zip Like "#########" Then LetterZip = Left$(zip, 5) & "-" & Right$(zip, 4)
End Function

'Resources > ISICs.
Public Function UXISICLoad() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXISICs
    f!lstISICs.RowSource = "SELECT ISICID,Title AS [ISIC],Nz(UIC,'') AS [UIC] FROM ISIC WHERE IsArchived=False ORDER BY Title"
    f!lstISICs.Requery
    UXISICLoad = UXISICNew()
    UXSortHeaders f
End Function
Public Function UXISICNew() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXISICs: isicID = 0: isicVersion = 0
    f!lstISICs = Null: f!txtISICName = Null: f!txtISICUIC = Null: f!txtISICAddress = Null
    f!lblStatus.Caption = "Enter a new ISIC, or choose one from the list to change it."
    UXISICNew = True
End Function
Public Function UXISICSelect() As Boolean
    Dim f As Access.Form, r As DAO.Recordset
    Set f = Forms!frmUXISICs
    If IsNull(f!lstISICs) Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ISIC WHERE ISICID=" & CLng(f!lstISICs), dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    isicID = r!ISICID: isicVersion = r!VersionCounter
    f!txtISICName = r!Title: f!txtISICUIC = r!UIC: f!txtISICAddress = r!Address: r.Close
    f!lblStatus.Caption = "Change the details, then Save ISIC."
    UXISICSelect = True
End Function
Public Function UXISICSave() As Boolean
    Dim f As Access.Form, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXISICs
    id = ISICSave(isicID, isicVersion, Nz(f!txtISICName, ""), Nz(f!txtISICUIC, ""), Nz(f!txtISICAddress, ""))
    Call UXISICLoad
    f!lstISICs = id: Call UXISICSelect
    f!lblStatus.Caption = "Saved. Choose this ISIC for a group's summary letter when creating its PDFs."
    UXISICSave = True: Exit Function
Failed:
    Forms!frmUXISICs!lblStatus.Caption = Err.Description
End Function
Public Function UXISICDelete() As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXISICs
    If isicID = 0 Then Err.Raise vbObjectError + 9905, "modISIC.UXISICDelete", "Choose an ISIC to delete."
    If MsgBox("Delete " & Nz(f!txtISICName, "this ISIC") & "? It will no longer be offered for summary letters.", vbYesNo + vbQuestion, "Delete ISIC") <> vbYes Then Exit Function
    ISICDelete isicID
    Call UXISICLoad
    f!lblStatus.Caption = "Deleted."
    UXISICDelete = True: Exit Function
Failed:
    Forms!frmUXISICs!lblStatus.Caption = Err.Description
End Function
