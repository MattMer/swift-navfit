Attribute VB_Name = "modUXReportReturn"
Option Explicit
Private Type ScrollInfo
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hwnd As LongPtr, ByVal bar As Long, ByRef info As ScrollInfo) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal message As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private captured As Boolean
Private topReport As Long
Private topPosition As Long
Private selectedReport As Long
Public Function UXReportScroll(Optional ByVal position As Long = -1) As Long
    Dim info As ScrollInfo, hwnd As LongPtr
    Forms!frmLibrary!lstReports.SetFocus
    hwnd = GetFocus(): info.cbSize = LenB(info): info.fMask = 23
    'A short/empty list has no scrollbar. It still supports opening reports.
    If GetScrollInfo(hwnd, 1, info) = 0 Then Exit Function
    If position >= 0 Then
        position = IIf(position > info.nMax, info.nMax, position)
        Call SendMessage(hwnd, &H115, CLngPtr(position) * 65536 + 4, 0)
        Call GetScrollInfo(hwnd, 1, info)
    End If
    UXReportScroll = info.nPos
End Function
Public Function UXCaptureReportList() As Boolean
    Dim f As Access.Form, row As Long, value As Variant
    Set f = Forms!frmLibrary
    selectedReport = Nz(f!lstReports, 0): topPosition = UXReportScroll()
    topReport = 0
    row = topPosition
    If row = 0 And f!lstReports.ColumnHeads Then row = 1
    If row < f!lstReports.ListCount Then
        value = f!lstReports.Column(0, row)
        If IsNumeric(value) Then topReport = CLng(value)
    End If
    captured = True: UXCaptureReportList = True
End Function
Public Function UXRestoreReportList() As Boolean
    Dim f As Access.Form, i As Long, position As Long, first As Long, value As Variant
    If Not captured Then UXRestoreReportList = True: Exit Function
    captured = False
    If Not CurrentProject.AllForms("frmLibrary").IsLoaded Then Exit Function
    Set f = Forms!frmLibrary
    Call RefreshLibrary
    f!lstReports = Null: position = topPosition
    If f!lstReports.ColumnHeads Then first = 1
    For i = first To f!lstReports.ListCount - 1
        value = f!lstReports.Column(0, i)
        If IsNumeric(value) Then
            If CLng(value) = selectedReport Then f!lstReports = selectedReport
            If topReport > 0 And CLng(value) = topReport Then
                position = i
                If i = first Then position = 0
            End If
        End If
    Next
    Call UXReportScroll(position)
    UXRestoreReportList = True
End Function
Public Function UXQueueReportReturn() As Boolean
    If captured And CurrentProject.AllForms("frmLibrary").IsLoaded Then Forms!frmLibrary.TimerInterval = 1
    UXQueueReportReturn = True
End Function
Public Function UXFinishReportReturn() As Boolean
    On Error GoTo Failed
    Forms!frmLibrary.TimerInterval = 0
    UXFinishReportReturn = UXRestoreReportList()
    Exit Function
Failed:
    Debug.Print "Reports return: " & Err.Description
End Function

