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
                LibraryViewSelect = "ReportID,FormKind,Title,MemberName,PeriodStart,PeriodEnd"
            Case "History"
                .ColumnCount = 6: .ColumnWidths = "0;1300;2700;2200;1400;1300"
                LibraryViewSelect = "ReportID,FormKind,Title,MemberName,ModifiedAt,CurrentRevisionID"
            Case Else
                .ColumnCount = 6: .ColumnWidths = "0;1300;2900;2700;1300;800"
                LibraryViewSelect = "ReportID,FormKind,Title,UXMemberLabel(ReportID) AS Member,UXReportState(ReportID) AS [Local status],IsArchived"
        End Select
    End With
    Forms!frmLibrary!cboView = preset
End Function
Public Function ChangeLibraryViewUI() As Boolean
    ChangeLibraryViewUI = SaveLibraryView(Nz(Forms!frmLibrary!cboView, "Drafting"))
    Call RefreshLibrary()
End Function
