Attribute VB_Name = "modLibraryViews"
Option Compare Database
Option Explicit
Public Function SaveLibraryView(ByVal preset As String) As Boolean
    Dim r As DAO.Recordset, pid As Long
    If preset <> "Drafting" And preset <> "Dates" And preset <> "History" Then Err.Raise vbObjectError + 1300, , "Unsupported library view."
    pid = ActiveContributorID()
    Set r = CurrentDb.OpenRecordset("SELECT * FROM LibraryPreference WHERE AuthorProfileID=" & pid, dbOpenDynaset)
    If r.EOF Then r.AddNew: r!AuthorProfileID = pid Else r.Edit
    r!ViewPreset = preset: r.Update: r.Close
    SaveLibraryView = True
End Function
Public Function LibraryViewSelect() As String
    Dim preset As String
    Call UXPrepareReportRows
    preset = Nz(DLookup("ViewPreset", "LibraryPreference", "AuthorProfileID=" & ActiveContributorID()), "Drafting")
    With Forms!frmLibrary!lstReports
        Select Case preset
            Case "Dates"
                .ColumnCount = 6: .ColumnWidths = "0;1300;2800;2600;1200;1200"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,PeriodStart AS [Period from],PeriodEnd AS [Period to]"
            Case "History"
                .ColumnCount = 6: .ColumnWidths = "0;1300;2700;2200;1400;1300"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,ModifiedAt AS [Last changed],UXReportRowCell(ReportID,4) AS [Saved revision]"
            Case Else
                .ColumnCount = 6: .ColumnWidths = "0;1300;2900;2700;1300;800"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,UXReportRowCell(ReportID,1) AS [Local status],IIf(IsArchived,'Archived','Active') AS [Archive state]"
        End Select
    End With
    LibraryViewSelect = LibraryViewSelect & ",UXReportRowCell(ReportID,2) AS ITA,PromotionRecommendation AS [Promotion recommendation],UXReportRowCell(ReportID,3) AS [Summary groups]"
    Forms!frmLibrary!lstReports.ColumnCount = 9
    Forms!frmLibrary!lstReports.ColumnWidths = "0;1300;2600;2500;2400;1500;900;2800;2600"
    Forms!frmLibrary!cboView = preset
End Function
Public Function ChangeLibraryViewUI() As Boolean
    ChangeLibraryViewUI = SaveLibraryView(Nz(Forms!frmLibrary!cboView, "Drafting"))
    Call RefreshLibrary()
End Function
