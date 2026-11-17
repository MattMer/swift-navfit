Attribute VB_Name = "modLibraryViews"
Option Compare Database
Option Explicit
Public Function SaveLibraryView(ByVal preset As String) As Boolean
    Dim r As DAO.Recordset, pid As Long
    If preset <> "Drafting" And preset <> "Dates" And preset <> "History" Then Err.Raise vbObjectError + 1300, "modLibraryViews.SaveLibraryView", "Unsupported library view."
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
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,RankRate AS [Rate],PeriodStart AS [Period from],PeriodEnd AS [Period to]"
            Case "History"
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,RankRate AS [Rate],ModifiedAt AS [Last changed],UXReportRowCell(ReportID,4) AS [Saved revision]"
            Case Else
                LibraryViewSelect = "ReportID,FormKind AS [Report type],Title,UXIdentity(OfficialName,MemberName) AS Member,RankRate AS [Rate],UXReportRowCell(ReportID,1) AS [Local status],IIf(IsArchived,'Archived','Active') AS [Archive state]"
        End Select
    End With
    LibraryViewSelect = LibraryViewSelect & ",UXReportRowCell(ReportID,2) AS ITA,PromotionRecommendation AS [Promotion recommendation],UXReportRowCell(ReportID,3) AS [Summary groups]"
    'Fit every column inside the list as it is now. Window resizing then scales them
    'in proportion, so the first load and a later refresh show the same layout.
    With Forms!frmLibrary!lstReports
        .ColumnCount = 10
        .ColumnWidths = LibraryFitWidths("0;1300;2500;2400;800;2300;1400;900;2700;2500", .Width - 300)
    End With
    Forms!frmLibrary!cboView = preset
End Function
' Scales column widths in proportion so they add up to room (twips). Hidden (zero)
' columns stay hidden.
Public Function LibraryFitWidths(ByVal design As String, ByVal room As Long) As String
    Dim parts As Variant, i As Long, total As Long, sizes As String
    parts = Split(design, ";")
    For i = 0 To UBound(parts): total = total + Val(parts(i)): Next
    For i = 0 To UBound(parts)
        If i > 0 Then sizes = sizes & ";"
        If total > 0 And room > 0 Then sizes = sizes & CLng(Val(parts(i)) * room / total) Else sizes = sizes & Val(parts(i))
    Next
    LibraryFitWidths = sizes
End Function
Public Function ChangeLibraryViewUI() As Boolean
    ChangeLibraryViewUI = SaveLibraryView(Nz(Forms!frmLibrary!cboView, "Drafting"))
    Call RefreshLibrary()
End Function
