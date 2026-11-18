Attribute VB_Name = "EvalmanHarness"
Option Compare Database
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6950, , message
    n = n + 1
End Sub
Private Function FindTitle(ByVal prefix As String) As Long
    Dim i As Long
    FindTitle = -1
    For i = 0 To EvalmanCount() - 1
        If Left$(EvalmanHeader(i, 1), Len(prefix)) = prefix Then FindTitle = i: Exit Function
    Next
End Function
Public Function TestEvalman() As String
    Dim i As Long, t As String, total As Long, hits As Collection, h As Variant, terms As Variant, f As Access.Form, before As String
    On Error GoTo Failed
    ' The whole instruction is present, in sections the reader can show.
    Check EvalmanCount() > 250 And EvalmanCount() < 700, "Unexpected section count " & EvalmanCount()
    Check InStr(EvalmanSource(), "1610.10H CH-4") > 0, "Source change missing: " & EvalmanSource()
    For i = 0 To EvalmanCount() - 1
        t = EvalmanText(i): total = total + Len(t)
        Check Len(t) < 60000, "Section too long for the reader: " & EvalmanHeader(i, 1)
        Check Len(EvalmanHeader(i, 1)) > 0 And EvalmanHeader(i, 0) >= 0 And EvalmanHeader(i, 0) <= 3, "Bad section header " & i
        Check InStr(t, ChrW$(65533)) = 0, "Unreadable character in " & EvalmanHeader(i, 1)
    Next
    Check total > 300000, "Text is incomplete: " & total & " characters"
    For Each h In Array("Chapter 1 - ", "1-4. Submission/Mailing Procedures", "Block 26: UIC", "Exhibit 1-1", "13-5. Prohibited Comments", "Chapter 19 - ")
        Check FindTitle(CStr(h)) >= 0, "Section missing: " & h
    Next
    Check InStr(EvalmanText(FindTitle("Block 26: UIC")), "same reporting senior UIC") > 0, "Block 26 text missing"
    ' Search: every word and quoted phrase must be present; matches point at the first term.
    terms = EvalmanTerms("  ""Early Promote""  Block ")
    Check UBound(terms) = 1 And terms(0) = "early promote" And terms(1) = "block", "Search terms parsed wrongly"
    Check UBound(EvalmanTerms("   ")) < 0, "Blank search produced terms"
    Set hits = EvalmanFind("""early promote""")
    Check hits.Count > 5, "Phrase search found too little: " & hits.Count
    For Each h In hits
        Check LCase$(Mid$(EvalmanText(CLng(h(0))), CLng(h(1)), CLng(h(2)))) = "early promote", "Match position wrong in " & EvalmanHeader(CLng(h(0)), 1)
    Next
    Check EvalmanFind("zzqxv").Count = 0, "Nonsense search matched"
    Set hits = EvalmanFind("delegated uic")
    For Each h In hits
        t = LCase$(EvalmanText(CLng(h(0))))
        Check InStr(t, "delegated") > 0 And InStr(t, "uic") > 0, "Section lacks a search word"
    Next
    Check hits.Count > 0, "Multi-word search found nothing"
    ' The reader: contents, results, stepping through matches and back to contents.
    Check UXOpenEvalman("""early promote"""), "Reader did not open"
    Set f = Forms!frmUXEvalman
    Check f!lstNav.ListCount > 0 And InStr(f!lblNav.Caption, "matches") > 0, "Results not listed"
    Check LCase$(f!txtReader.SelText) = "early promote", "Match not highlighted: " & f!txtReader.SelText
    Check f!lblMatch.Caption = "Match 1 of " & EvalmanFind("""early promote""").Count, "Match counter wrong: " & f!lblMatch.Caption
    Check UXEvalmanNext() And InStr(f!lblMatch.Caption, "Match 2 of") = 1 And LCase$(f!txtReader.SelText) = "early promote", "Next match failed"
    Check UXEvalmanPrevious() And InStr(f!lblMatch.Caption, "Match 1 of") = 1, "Previous match failed"
    Check UXEvalmanPrevious() And InStr(f!lblMatch.Caption, "Match " & EvalmanFind("""early promote""").Count & " of") = 1, "Previous did not wrap to the last match"
    f!txtSearch = "zzqxv": Check UXEvalmanSearch() And f!lstNav.ListCount = 0 And f!lblNav.Caption = "No matches", "Empty search result wrong"
    Check UXEvalmanShowContents() And f!lstNav.ListCount = EvalmanCount() And f!lblNav.Caption = "Contents", "Contents not restored"
    i = FindTitle("3-9. Special Reports")
    f!lstNav = CStr(i): Check UXEvalmanSelect() And UXEvalmanShown() = i, "Contents selection did not open the section"
    Check Left$(f!lblReader.Caption, 20) = "3-9. Special Reports" And InStr(f!txtReader, "Special reports are considered") > 0, "Section text not shown"
    f!txtSearch = Null: Check UXEvalmanSearch() And f!lblNav.Caption = "Contents", "Clearing the search did not show contents"
    DoCmd.Close acForm, "frmUXEvalman", acSaveNo
    ' Resources opens it.
    DoCmd.OpenForm "frmUXResources"
    Check Forms!frmUXResources!cmdEvalman.Visible And Forms!frmUXResources!cmdEvalman.OnClick = "=UXOpenEvalman()", "Resources link missing"
    DoCmd.Close acForm, "frmUXResources", acSaveNo
    TestEvalman = "PASS: " & n & " EVALMAN text, search and reader checks": Exit Function
Failed:
    TestEvalman = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmUXEvalman", acSaveNo
End Function
