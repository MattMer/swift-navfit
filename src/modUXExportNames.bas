Attribute VB_Name = "modUXExportNames"
Option Explicit
Private Function SafeRank(ByVal value As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(value)
        ch = UCase$(Mid$(value, i, 1))
        If ch Like "[A-Z0-9_-]" Then SafeRank = SafeRank & ch
    Next
    If Len(SafeRank) = 0 Then SafeRank = "UNSPECIFIED"
End Function
Public Function UXExportName(ByVal groupID As Long, ByVal reportID As Long) As String
    Dim rank As String, ending As Variant, dates As Object, r As DAO.Recordset, db As DAO.Database, key As Variant
    If groupID > 0 Then
        rank = Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID), "")
        Set dates = CreateObject("Scripting.Dictionary"): Set db = CurrentDb
        Set r = db.OpenRecordset("SELECT r.PeriodEnd FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID, dbOpenSnapshot)
        Do While Not r.EOF
            If IsNull(r!PeriodEnd) Then ending = "UNSPECIFIED" Else ending = Format$(r!PeriodEnd, "yyyy-mm-dd")
            dates(CStr(ending)) = True: r.MoveNext
        Loop
        r.Close
        ending = "MIXED"
        If dates.Count = 1 Then
            For Each key In dates.Keys: ending = CStr(key): Next
        End If
    Else
        rank = Nz(DLookup("RankRate", "Report", "ReportID=" & reportID), "")
        ending = DLookup("PeriodEnd", "Report", "ReportID=" & reportID)
        If IsNull(ending) Then Err.Raise vbObjectError + 6910, , "Set the saved report's ending date before exporting."
        ending = Format$(ending, "yyyy-mm-dd")
    End If
    UXExportName = "Swift_NAVFIT_Export_" & SafeRank(rank) & "_" & CStr(ending) & ".accdb"
End Function
Public Function UXExportDestination(ByVal chosen As String, ByVal groupID As Long, ByVal reportID As Long) As String
    Dim fs As Object, folder As String, name As String, stem As String, path As String, n As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = chosen
    If LCase$(Right$(chosen, 6)) = ".accdb" Then folder = fs.GetParentFolderName(chosen)
    If Not fs.FolderExists(folder) Then Err.Raise vbObjectError + 6911, , "Choose an existing export directory."
    name = UXExportName(groupID, reportID): stem = Left$(name, Len(name) - 6): path = fs.BuildPath(folder, name)
    Do While fs.FileExists(path) Or fs.FileExists(path & ".swift-baseline.accdb") Or fs.FolderExists(path)
        n = n + 1
        If n > 10000 Then Err.Raise vbObjectError + 6912, , "Choose another export directory; too many matching export names exist."
        path = fs.BuildPath(folder, stem & "_" & n & ".accdb")
    Loop
    UXExportDestination = path
End Function
