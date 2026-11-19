Attribute VB_Name = "modUXEvalman"
Option Compare Database
Option Explicit
' EVALMAN reader: the full text of BUPERSINST 1610.10H (modEvalmanText), its contents,
' and search. A search matches sections that contain every word or "quoted phrase";
' each occurrence of the first term is a match the reader can step through.
Private lowers() As String, ready As Boolean
Private hits As Collection, hitIndex As Long, terms As Variant
Private listMode As String, shown As Long
Private Const MAX_ROWS As Long = 150

Private Sub EnsureText()
    Dim i As Long
    If ready Then Exit Sub
    ReDim lowers(0 To EvalmanCount() - 1)
    For i = 0 To EvalmanCount() - 1
        lowers(i) = LCase$(EvalmanText(i))
    Next
    ready = True
End Sub

' Words and "quoted phrases" from a search, lower case, in the order typed.
Public Function EvalmanTerms(ByVal query As String) As Variant
    Dim out As String, part As Variant, i As Long, quoted As Variant
    query = LCase$(Trim$(query))
    quoted = Split(query, """")
    For i = 0 To UBound(quoted)
        If i Mod 2 = 1 Then
            If Len(Trim$(quoted(i))) > 0 Then out = out & Chr$(1) & Trim$(quoted(i))
        Else
            For Each part In Split(quoted(i), " ")
                If Len(part) > 0 Then out = out & Chr$(1) & part
            Next
        End If
    Next
    If Len(out) = 0 Then EvalmanTerms = Array() Else EvalmanTerms = Split(Mid$(out, 2), Chr$(1))
End Function

' Every occurrence of the first term in sections that contain all the terms:
' items are Array(section, start, length), start 1-based in EvalmanText(section).
Public Function EvalmanFind(ByVal query As String) As Collection
    Dim found As New Collection, i As Long, k As Long, p As Long, all As Boolean
    terms = EvalmanTerms(query)
    Set EvalmanFind = found
    If UBound(terms) < 0 Then Exit Function
    EnsureText
    For i = 0 To EvalmanCount() - 1
        all = True
        For k = 0 To UBound(terms)
            If InStr(1, lowers(i), terms(k), vbBinaryCompare) = 0 Then all = False: Exit For
        Next
        If all Then
            p = InStr(1, lowers(i), terms(0), vbBinaryCompare)
            Do While p > 0
                found.Add Array(i, p, Len(terms(0)))
                p = InStr(p + Len(terms(0)), lowers(i), terms(0), vbBinaryCompare)
            Loop
        End If
    Next
End Function

Private Function Quoted(ByVal value As String) As String
    Quoted = """" & Replace(value, """", """""") & """"
End Function

Public Function UXOpenEvalman(Optional ByVal query As String = "") As Boolean
    DoCmd.OpenForm "frmUXEvalman"
    If Len(query) > 0 Then
        Forms!frmUXEvalman!txtSearch = query
        Call UXEvalmanSearch
    End If
    UXOpenEvalman = True
End Function

Public Function UXLoadEvalman() As Boolean
    With Forms!frmUXEvalman
        !lblSource.Caption = EvalmanSource() & ". Text from the official PDF, which is the authoritative copy."
        !txtReader = Null: !lblReader.Caption = "Choose a section"
    End With
    Set hits = New Collection: hitIndex = 0: shown = -1
    UXEvalmanShowContents
    UXLoadEvalman = True
End Function

Public Function UXEvalmanShowContents() As Boolean
    Dim i As Long, rows As String
    For i = 0 To EvalmanCount() - 1
        rows = rows & ";" & i & ";" & Quoted(String$(CLng(EvalmanHeader(i, 0)) * 3, ChrW$(160)) & EvalmanHeader(i, 1))
    Next
    With Forms!frmUXEvalman
        !lstNav.RowSource = Mid$(rows, 2)
        !lblNav.Caption = "Contents"
        !lblMatch.Caption = ""
        !cmdPrev.Enabled = False: !cmdNext.Enabled = False
    End With
    listMode = "Contents"
    UXEvalmanShowContents = True
End Function

Public Function UXEvalmanSearch() As Boolean
    Dim f As Access.Form, h As Variant, rows As String, n As Long, counts As Object, s As Variant
    Set f = Forms!frmUXEvalman
    If Len(Trim$(Nz(f!txtSearch, ""))) = 0 Then UXEvalmanSearch = UXEvalmanShowContents(): Exit Function
    Set hits = EvalmanFind(Nz(f!txtSearch, ""))
    Set counts = CreateObject("Scripting.Dictionary")
    For Each h In hits
        If Not counts.Exists(h(0)) Then counts.Add h(0), 0
        counts(h(0)) = counts(h(0)) + 1
    Next
    For Each s In counts.Keys
        n = n + 1
        If n > MAX_ROWS Then Exit For
        rows = rows & ";" & s & ";" & Quoted(EvalmanHeader(CLng(s), 1) & " (" & counts(s) & ")")
    Next
    f!lstNav.RowSource = Mid$(rows, 2)
    listMode = "Results": hitIndex = 0
    If hits.Count = 0 Then
        f!lblNav.Caption = "No matches"
        f!lblMatch.Caption = "Nothing found. Try fewer words."
    Else
        f!lblNav.Caption = hits.Count & " matches in " & counts.Count & " sections" & IIf(counts.Count > MAX_ROWS, " (first " & MAX_ROWS & " listed)", "")
        GoToHit 1
    End If
    f!cmdPrev.Enabled = hits.Count > 1: f!cmdNext.Enabled = hits.Count > 1
    UXEvalmanSearch = True
End Function

Private Sub ShowSection(ByVal section As Long)
    With Forms!frmUXEvalman
        If shown <> section Then
            !txtReader = EvalmanText(section)
            !lblReader.Caption = EvalmanHeader(section, 1) & IIf(Len(EvalmanHeader(section, 2)) > 0, ", page " & EvalmanHeader(section, 2), "")
            shown = section
        End If
    End With
End Sub

Private Sub GoToHit(ByVal index As Long)
    Dim h As Variant
    If hits Is Nothing Then Exit Sub
    If hits.Count = 0 Then Exit Sub
    If index < 1 Then index = hits.Count
    If index > hits.Count Then index = 1
    hitIndex = index: h = hits(index)
    ShowSection CLng(h(0))
    With Forms!frmUXEvalman
        If listMode = "Results" Then .lstNav = CStr(h(0))
        !txtReader.SetFocus
        !txtReader.SelStart = CLng(h(1)) - 1
        !txtReader.SelLength = CLng(h(2))
        !lblMatch.Caption = "Match " & index & " of " & hits.Count
    End With
End Sub

Public Function UXEvalmanSelect() As Boolean
    Dim section As Long, i As Long
    section = Val(Nz(Forms!frmUXEvalman!lstNav, -1))
    If section < 0 Then Exit Function
    If listMode = "Results" Then
        For i = 1 To hits.Count
            If hits(i)(0) = section Then GoToHit i: Exit For
        Next
    Else
        ShowSection section
        Forms!frmUXEvalman!txtReader.SetFocus
        Forms!frmUXEvalman!txtReader.SelStart = 0
    End If
    UXEvalmanSelect = True
End Function

Public Function UXEvalmanNext() As Boolean
    GoToHit hitIndex + 1: UXEvalmanNext = True
End Function

Public Function UXEvalmanPrevious() As Boolean
    GoToHit hitIndex - 1: UXEvalmanPrevious = True
End Function

' Open the reader at the top of one section (a finding's paragraph or a field's block),
' with the contents listed and that section selected.
Public Function UXOpenEvalmanSection(ByVal section As Long) As Boolean
    If section < 0 Or section >= EvalmanCount() Then Exit Function
    DoCmd.OpenForm "frmUXEvalman"
    With Forms!frmUXEvalman
        If listMode <> "Contents" Or Len(Nz(!txtSearch, "")) > 0 Then !txtSearch = Null: UXEvalmanShowContents
        !lstNav = CStr(section)
        ShowSection section
        !txtReader.SetFocus
        !txtReader.SelStart = 0: !txtReader.SelLength = 0
        !lblMatch.Caption = ""
    End With
    UXOpenEvalmanSection = True
End Function

' The section a reader is showing, for tests and links from other windows.
Public Function UXEvalmanShown() As Long
    UXEvalmanShown = shown
End Function
