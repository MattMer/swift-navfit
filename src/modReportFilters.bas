Attribute VB_Name = "modReportFilters"
Option Compare Database
Option Explicit
Private selectedFilter As Long, selectedVersion As Long, selectedOwner As String, selectedName As String

'Search keeps the stored spelling intact. Commas separate words; apostrophes
'and hyphens may be omitted, and every query word can occur in any order.
Private Function ReportSearchWords(ByVal value As String) As String
    Dim i As Long, ch As String, code As Long
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch)
        Select Case code
            Case 39, 45, 8217, 8208, 8209, 8211, 8212
            Case Else
                If ch Like "[A-Za-z0-9]" Or code < 0 Or code >= 192 Then
                    ReportSearchWords = ReportSearchWords & UCase$(ch)
                Else
                    ReportSearchWords = ReportSearchWords & " "
                End If
        End Select
    Next
End Function
Public Function ReportSearchMatches(ByVal query As String, ByVal title As Variant, ByVal member As Variant, ByVal official As Variant) As Boolean
    Dim word As Variant, a As String, b As String, c As String
    a = Replace(ReportSearchWords(Nz(title, "")), " ", "")
    b = Replace(ReportSearchWords(Nz(member, "")), " ", "")
    c = Replace(ReportSearchWords(Nz(official, "")), " ", "")
    For Each word In Split(ReportSearchWords(Left$(query, 255)), " ")
        If Len(word) > 0 Then
            If InStr(1, a, CStr(word), vbTextCompare) = 0 And InStr(1, b, CStr(word), vbTextCompare) = 0 And InStr(1, c, CStr(word), vbTextCompare) = 0 Then Exit Function
        End If
    Next
    ReportSearchMatches = True
End Function
Private Function ReportFilterDate(ByVal value As Variant) As Variant
    ReportFilterDate = Null
    If IsNull(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If Not IsDate(value) Then Err.Raise vbObjectError + 9820, "modReportFilters", "Enter a valid period date or leave it blank."
    ReportFilterDate = DateValue(CDate(value))
End Function
Private Function ReportFilterSQLDate(ByVal value As Date) As String
    ReportFilterSQLDate = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function
Public Function ReportPeriodWhere(ByVal fromDate As Variant, ByVal throughDate As Variant) As String
    fromDate = ReportFilterDate(fromDate): throughDate = ReportFilterDate(throughDate)
    If Not IsNull(fromDate) And Not IsNull(throughDate) Then
        If fromDate > throughDate Then Err.Raise vbObjectError + 9821, "modReportFilters", "The period from date must be on or before the through date."
    End If
    If Not IsNull(fromDate) Then ReportPeriodWhere = " AND PeriodEnd>=" & ReportFilterSQLDate(CDate(fromDate))
    If Not IsNull(throughDate) Then ReportPeriodWhere = ReportPeriodWhere & " AND PeriodStart<=" & ReportFilterSQLDate(CDate(throughDate))
End Function
Public Function ReportFilterCriteria() As Object
    Dim result As Object, f As Access.Form, groupID As Long
    Set result = CreateObject("Scripting.Dictionary"): Set f = Forms!frmLibrary
    result.Add "SearchText", Left$(Nz(f!txtSearch, ""), 255)
    result.Add "PeriodFrom", ReportFilterDate(f!txtPeriodFrom): result.Add "PeriodThrough", ReportFilterDate(f!txtPeriodThrough)
    result.Add "FormKind", Nz(f!uxTypeFilter, "All"): result.Add "ReportState", Nz(f!uxStatusFilter, "All")
    result.Add "RankFilter", Nz(f!cboRankFilter, "All ranks/rates"): result.Add "ArchiveMode", Nz(f!cboArchive, "Active")
    groupID = Nz(f!lstFolders, 0)
    result.Add "GroupScope", IIf(groupID = -1, "Ungrouped", IIf(groupID > 0, "Group", "All"))
    result.Add "GroupGuid", Null
    If groupID > 0 Then result("GroupGuid") = MajorScalar("SELECT GroupGuid FROM SummaryGroup WHERE GroupID=" & groupID)
    Set ReportFilterCriteria = result
End Function
Private Sub CheckReportFilter(ByVal values As Object)
    Dim unused As String, key As Variant
    If values.Count <> 9 Then Err.Raise vbObjectError + 9822, "modReportFilters", "The saved filter has unexpected fields."
    For Each key In Array("SearchText", "PeriodFrom", "PeriodThrough", "FormKind", "ReportState", "RankFilter", "ArchiveMode", "GroupScope", "GroupGuid")
        If Not values.Exists(CStr(key)) Then Err.Raise vbObjectError + 9822, "modReportFilters", "The saved filter is incomplete."
    Next
    unused = ReportPeriodWhere(values("PeriodFrom"), values("PeriodThrough"))
    If Not ReportFilterChoice(values("FormKind"), "All|EVAL|CHIEFEVAL|FITREP") Then Err.Raise vbObjectError + 9822, "modReportFilters", "Unknown report type in the saved filter."
    If Not ReportFilterChoice(values("ReportState"), "All|Draft|In review|Review needs restart|Local approval|Signed|Returned") Then Err.Raise vbObjectError + 9822, "modReportFilters", "Unknown report status in the saved filter."
    If Not ReportFilterChoice(values("ArchiveMode"), "Active|Archived|All") Then Err.Raise vbObjectError + 9822, "modReportFilters", "Unknown archive mode in the saved filter."
    If values("GroupScope") <> "All" And values("GroupScope") <> "Ungrouped" And values("GroupScope") <> "Group" Then Err.Raise vbObjectError + 9822, "modReportFilters", "Unknown group scope in the saved filter."
    If values("GroupScope") = "Group" And Len(Nz(values("GroupGuid"), "")) = 0 Then Err.Raise vbObjectError + 9822, "modReportFilters", "The saved filter's group is unavailable."
    If Len(Nz(values("SearchText"), "")) > 255 Or Len(Nz(values("RankFilter"), "")) > 40 Then Err.Raise vbObjectError + 9822, "modReportFilters", "The saved filter contains an overlong value."
End Sub
Private Function ReportFilterChoice(ByVal value As Variant, ByVal choices As String) As Boolean
    Dim item As Variant
    If IsNull(value) Then Exit Function
    For Each item In Split(choices, "|")
        If CStr(value) = CStr(item) Then ReportFilterChoice = True: Exit Function
    Next
End Function
Private Function ReportFilterOwner() As String
    ReportFilterOwner = CStr(MajorScalar("SELECT ProfileGuid FROM UserProfile WHERE UserProfileID=" & ActiveContributorID()))
End Function
Public Function SaveReportFilter(ByVal filterID As Long, ByVal expectedVersion As Long, ByVal filterName As String, ByVal values As Object) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, owner As String, key As Variant
    owner = ReportFilterOwner(): filterName = Trim$(filterName)
    If Len(filterName) = 0 Or Len(filterName) > 80 Then Err.Raise vbObjectError + 9823, "modReportFilters", "Enter a filter name of 1 to 80 characters."
    CheckReportFilter values
    If DCount("*", "SavedReportFilter", "OwnerProfileGuid=" & QuoteSQL(owner) & " AND FilterName=" & QuoteSQL(filterName) & " AND FilterID<>" & filterID) > 0 Then Err.Raise vbObjectError + 9823, "modReportFilters", "That filter name already exists for this contributor."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM SavedReportFilter WHERE FilterID=" & filterID & " AND OwnerProfileGuid=" & QuoteSQL(owner), dbOpenDynaset)
    If filterID = 0 Then
        r.AddNew: r!FilterGuid = NewIdentity(): r!OwnerProfileGuid = owner: r!VersionCounter = 1
    Else
        If r.EOF Then Err.Raise vbObjectError + 9824, "modReportFilters", "This saved filter is unavailable for the current contributor."
        If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 9825, "modReportFilters", "The saved filter changed. Select it again before saving."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!FilterName = filterName
    For Each key In Array("SearchText", "PeriodFrom", "PeriodThrough", "FormKind", "ReportState", "RankFilter", "ArchiveMode", "GroupScope", "GroupGuid"): r.Fields(CStr(key)).Value = values(key): Next
    r!ModifiedAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveReportFilter = r!FilterID: r.Close
    w.CommitTrans: Exit Function
Failed:
    Dim primary As Object
    Set primary = SwiftFailure("modReportFilters", "SaveReportFilter", "Save", Err.Number, Err.Description, Err.Source, "Active")
    SwiftRollback w, primary: If Not r Is Nothing Then SwiftClose r, primary, "r"
    On Error GoTo 0: SwiftRaise primary
End Function
Public Function DeleteReportFilter(ByVal filterID As Long, ByVal expectedVersion As Long) As Boolean
    Dim q As DAO.QueryDef
    Set q = CurrentDb.CreateQueryDef("", "DELETE FROM SavedReportFilter WHERE FilterID=" & filterID & " AND VersionCounter=" & expectedVersion & " AND OwnerProfileGuid=" & QuoteSQL(ReportFilterOwner()))
    q.Execute dbFailOnError
    If q.RecordsAffected <> 1 Then q.Close: Err.Raise vbObjectError + 9825, "modReportFilters", "The saved filter changed or belongs to another contributor. Select it again."
    q.Close: DeleteReportFilter = True
End Function
Public Function UXReportFilterChoices() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmLibrary
    f!cboSavedFilter.RowSource = "SELECT FilterID,FilterName,VersionCounter FROM SavedReportFilter WHERE OwnerProfileGuid=" & QuoteSQL(ReportFilterOwner()) & " ORDER BY FilterName"
    f!cboSavedFilter.Requery: f!cboSavedFilter = Null
    selectedFilter = 0: selectedVersion = 0: selectedOwner = "": selectedName = ""
    UXReportFilterChoices = True
End Function
Public Function UXLoadReportFilter() As Boolean
    Dim r As DAO.Recordset, f As Access.Form, values As Object, key As Variant, groupID As Long, explanation As String
    On Error GoTo Failed
    Set f = Forms!frmLibrary
    If IsNull(f!cboSavedFilter) Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM SavedReportFilter WHERE FilterID=" & CLng(f!cboSavedFilter) & " AND OwnerProfileGuid=" & QuoteSQL(ReportFilterOwner()), dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 9824, "modReportFilters", "This saved filter is unavailable for the current contributor."
    selectedFilter = r!FilterID: selectedVersion = r!VersionCounter: selectedOwner = r!OwnerProfileGuid: selectedName = r!FilterName
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In Array("SearchText", "PeriodFrom", "PeriodThrough", "FormKind", "ReportState", "RankFilter", "ArchiveMode", "GroupScope", "GroupGuid"): values.Add CStr(key), r.Fields(CStr(key)).Value: Next
    r.Close: Set r = Nothing: CheckReportFilter values
    If values("GroupScope") = "Ungrouped" Then groupID = -1
    If values("GroupScope") = "Group" Then
        groupID = Nz(MajorScalar("SELECT GroupID FROM SummaryGroup WHERE IsArchived=False AND GroupGuid=" & QuoteSQL(values("GroupGuid"))), 0)
        If groupID = 0 Then Err.Raise vbObjectError + 9824, "modReportFilters", "The saved filter's group is unavailable. Choose another filter or clear filters."
    End If
    f!txtSearch = values("SearchText"): f!txtPeriodFrom = values("PeriodFrom"): f!txtPeriodThrough = values("PeriodThrough")
    f!uxTypeFilter = values("FormKind"): f!uxStatusFilter = values("ReportState"): f!cboRankFilter = values("RankFilter"): f!cboArchive = values("ArchiveMode")
    f!lstFolders = groupID: f!uxFolderFilter = groupID
    UXLoadReportFilter = RefreshLibrary(): Exit Function
Failed:
    explanation = Err.Description
    If Not r Is Nothing Then r.Close
    Forms!frmLibrary!lblCount.Caption = explanation
End Function
Private Sub RequireReportFilterSelection(ByVal id As Long)
    If selectedFilter <> id Or selectedOwner <> ReportFilterOwner() Or selectedVersion < 1 Then Err.Raise vbObjectError + 9825, "modReportFilters", "Select the saved filter again before changing it."
End Sub
Public Function UXSaveReportFilter(Optional ByVal name As String = "", Optional ByVal confirmed As Boolean = False) As Boolean
    Dim f As Access.Form, id As Long, version As Long, existing As String
    On Error GoTo Failed
    Set f = Forms!frmLibrary
    If Not IsNull(f!cboSavedFilter) Then
        RequireReportFilterSelection CLng(f!cboSavedFilter): existing = selectedName
    End If
    If name = "" Then name = InputBox("Name this filter set. A new name saves a separate set.", "Save report filter", existing)
    If Len(Trim$(name)) = 0 Then Exit Function
    If StrComp(Trim$(name), existing, vbTextCompare) = 0 And existing <> "" Then
        id = CLng(f!cboSavedFilter): version = selectedVersion
        If Not confirmed Then
            If MsgBox("Replace the selected saved filter with these settings?", vbYesNo + vbQuestion + vbDefaultButton2, "Save report filter") <> vbYes Then Exit Function
        End If
    End If
    id = SaveReportFilter(id, version, name, ReportFilterCriteria())
    Call UXReportFilterChoices: f!cboSavedFilter = id: f!lblCount.Caption = "Filter saved for the current contributor."
    selectedFilter = id: selectedVersion = version + 1: selectedOwner = ReportFilterOwner(): selectedName = Trim$(name)
    UXSaveReportFilter = True: Exit Function
Failed:
    Forms!frmLibrary!lblCount.Caption = Err.Description
End Function
Public Function UXDeleteReportFilter(Optional ByVal confirmed As Boolean = False) As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmLibrary
    If IsNull(f!cboSavedFilter) Then f!lblCount.Caption = "Select a saved filter to delete.": Exit Function
    RequireReportFilterSelection CLng(f!cboSavedFilter)
    If Not confirmed Then
        If MsgBox("Delete the selected saved filter? Reports are unchanged.", vbYesNo + vbQuestion + vbDefaultButton2, "Delete report filter") <> vbYes Then Exit Function
    End If
    UXDeleteReportFilter = DeleteReportFilter(CLng(f!cboSavedFilter), selectedVersion)
    Call UXReportFilterChoices: f!lblCount.Caption = "Saved filter deleted; the current report filters remain."
    Exit Function
Failed:
    Forms!frmLibrary!lblCount.Caption = Err.Description
End Function
