Attribute VB_Name = "modLibraryViews"
Option Compare Database
Option Explicit
Public Function SaveLibraryView(ByVal preset As String) As Boolean
    Dim r As DAO.Recordset, pid As Long
    If preset <> "Drafting" And preset <> "Dates" And preset <> "History" Then Err.Raise vbObjectError + 1300, , "Unsupported library view."
    pid = DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1")
    Set r = CurrentDb.OpenRecordset("SELECT * FROM LibraryPreference WHERE AuthorProfileID=" & pid, dbOpenDynaset)
    If r.EOF Then r.AddNew: r!AuthorProfileID = pid Else r.Edit
    r!ViewPreset = preset: r.Update: r.Close
    SaveLibraryView = True
End Function
Public Function LibraryViewSelect() As String
    Dim preset As String
    preset = Nz(DLookup("ViewPreset", "LibraryPreference", "AuthorProfileID=" & Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)), "Drafting")
    With Forms!frmLibrary!lstReports
        Select Case preset
            Case "Dates"
                .ColumnCount = 6: .ColumnWidths = "0;1300;2800;2600;1200;1200"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXMemberLabel(ReportID) AS Member,PeriodStart AS [Period from],PeriodEnd AS [Period to]"
            Case "History"
                .ColumnCount = 6: .ColumnWidths = "0;1300;2700;2200;1400;1300"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXMemberLabel(ReportID) AS Member,ModifiedAt AS [Last changed],DLookup('RevisionNumber','ReportRevision','RevisionID=' & CurrentRevisionID) AS [Saved revision]"
            Case Else
                .ColumnCount = 6: .ColumnWidths = "0;1300;2900;2700;1300;800"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXMemberLabel(ReportID) AS Member,UXReportState(ReportID) AS [Local status],IIf(IsArchived,'Archived','Active') AS [Archive state]"
        End Select
    End With
    LibraryViewSelect = LibraryViewSelect & ",UXTableITA(ReportID) AS ITA,PromotionRecommendation AS [Promotion recommendation],UXTableGroups(ReportID) AS [Summary groups]"
    Forms!frmLibrary!lstReports.ColumnCount = 9
    Forms!frmLibrary!lstReports.ColumnWidths = "0;900;1700;1800;1250;1050;700;1500;1700"
    Forms!frmLibrary!cboView = preset
End Function
Public Function ChangeLibraryViewUI() As Boolean
    ChangeLibraryViewUI = SaveLibraryView(Nz(Forms!frmLibrary!cboView, "Drafting"))
    Call RefreshLibrary()
End Function
