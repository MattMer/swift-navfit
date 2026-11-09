Attribute VB_Name = "modRestrictedContent"
Option Compare Binary
Option Explicit
'Restricted-content check: possible prohibited or sensitive comments in narrative blocks,
'from the reviewed catalogue in modRestrictedRules (src/restricted-content.json).
'Findings are review prompts. A finding can be dismissed for one report, or its phrase
'allowed for every report; both are recorded and can be undone.
'
'Matching is on whole lower-case words (letters and digits; a trailing 's is dropped),
'so "LOC" never matches "LOCATION". A term ending in * matches a word prefix. A pattern
'with near terms counts only when one sits within "within" words; an unless term within
'three words of the match (or of its near term) suppresses it. Plain VBA string work: no
'regular-expression or scripting objects.
Private Const UNLESS_REACH As Long = 3
Private Const CACHE_LIMIT As Long = 64
Private compiled As Boolean
Private ruleRows() As Variant, ruleIndexByID As Object
Private groupRule() As Long, groupTerms() As Variant, groupNear() As Variant, groupWithin() As Long, groupUnless() As Variant, groupCount As Long
Private exceptionsLoaded As Boolean, exceptionTerms As Collection
Private results As Object, resultOrder As Collection
Private dismissedReport As Long, dismissed As Object

'--- Catalogue ---------------------------------------------------------------------------
Private Function SplitTerms(ByVal joined As String) As Variant
    Dim parts As Variant, i As Long, out() As Variant
    If Len(joined) = 0 Then SplitTerms = Array(): Exit Function
    parts = Split(joined, "~"): ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts): out(i) = Split(CStr(parts(i)), " "): Next
    SplitTerms = out
End Function
Private Sub Compile()
    Dim i As Long, row As Variant
    If compiled Then Exit Sub
    ReDim ruleRows(0 To RestrictedRuleCount() - 1)
    Set ruleIndexByID = CreateObject("Scripting.Dictionary")
    For i = 0 To RestrictedRuleCount() - 1
        ruleRows(i) = RestrictedRule(i): ruleIndexByID(CStr(ruleRows(i)(0))) = i
    Next
    groupCount = RestrictedMatchCount()
    ReDim groupRule(0 To groupCount - 1): ReDim groupTerms(0 To groupCount - 1): ReDim groupNear(0 To groupCount - 1)
    ReDim groupWithin(0 To groupCount - 1): ReDim groupUnless(0 To groupCount - 1)
    For i = 0 To groupCount - 1
        row = RestrictedMatch(i)
        groupRule(i) = CLng(row(0)): groupWithin(i) = CLng(row(3))
        If ruleRows(groupRule(i))(7) = "marking" Then groupTerms(i) = Split(CStr(row(1)), "~") Else groupTerms(i) = SplitTerms(CStr(row(1)))
        groupNear(i) = SplitTerms(CStr(row(2))): groupUnless(i) = SplitTerms(CStr(row(4)))
    Next
    compiled = True
End Sub
Public Function RestrictedRuleID(ByVal index As Long) As String
    Call Compile: RestrictedRuleID = CStr(ruleRows(index)(0))
End Function
Public Function RestrictedRuleTitle(ByVal ruleID As String) As String
    Compile
    If ruleIndexByID.Exists(ruleID) Then RestrictedRuleTitle = CStr(ruleRows(ruleIndexByID(ruleID))(1))
End Function
'The citation shown with a finding.
Public Function RestrictedReference(ByVal ruleID As String) As String
    Dim row As Variant
    Compile
    If Not ruleIndexByID.Exists(ruleID) Then Exit Function
    row = ruleRows(ruleIndexByID(ruleID))
    If Val(row(4)) > 0 Then
        RestrictedReference = RestrictedCatalogueSource() & ", " & row(2) & ", printed page " & row(3) & " (PDF page " & row(4) & ")."
    Else
        RestrictedReference = row(2) & "."
    End If
    RestrictedReference = RestrictedReference & vbCrLf & "Excerpt: " & Chr$(34) & row(5) & Chr$(34) & vbCrLf & _
        "Swift checks wording only and cannot tell whether, for example, a proceeding has concluded. Dismiss the finding for this report, or allow the phrase for every report, when the comment is acceptable." & vbCrLf & _
        "Rule catalogue SHA-256: " & RestrictedCatalogueHash()
End Function

'--- Words -------------------------------------------------------------------------------
'Lower-case words with their position and length in the original text.
Private Function Tokenize(ByVal text As String, ByRef starts() As Long, ByRef lengths() As Long) As String()
    Dim words() As String, count As Long, i As Long, c As String, code As Long, word As String, first As Long
    ReDim words(0 To 15): ReDim starts(0 To 15): ReDim lengths(0 To 15)
    For i = 1 To Len(text) + 1
        If i <= Len(text) Then c = Mid$(text, i, 1) Else c = " "
        code = AscW(c)
        If code = 8217 Or code = 8216 Then c = "'": code = 39
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code > 191 Or (code = 39 And Len(word) > 0) Then
            If Len(word) = 0 Then first = i
            word = word & LCase$(c)
        ElseIf Len(word) > 0 Then
            If Right$(word, 2) = "'s" Then word = Left$(word, Len(word) - 2)
            word = Replace(word, "'", "")
            If Len(word) > 0 Then
                If count > UBound(words) Then ReDim Preserve words(0 To count * 2): ReDim Preserve starts(0 To count * 2): ReDim Preserve lengths(0 To count * 2)
                words(count) = word: starts(count) = first: lengths(count) = i - first: count = count + 1
            End If
            word = ""
        End If
    Next
    If count = 0 Then
        ReDim words(0 To 0): words(0) = "": starts(0) = 0: lengths(0) = 0: Tokenize = words: Exit Function
    End If
    ReDim Preserve words(0 To count - 1): ReDim Preserve starts(0 To count - 1): ReDim Preserve lengths(0 To count - 1)
    Tokenize = words
End Function
Private Function WordMatches(ByVal pattern As String, ByVal word As String) As Boolean
    If Right$(pattern, 1) = "*" Then
        WordMatches = (Left$(word, Len(pattern) - 1) = Left$(pattern, Len(pattern) - 1))
    Else
        WordMatches = (word = pattern)
    End If
End Function
'Word positions by word, built once per text.
Private Function WordIndex(ByRef words() As String) As Object
    Dim i As Long, index As Object
    Set index = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            If Not index.Exists(words(i)) Then index.Add words(i), New Collection
            index(words(i)).Add i
        End If
    Next
    Set WordIndex = index
End Function
'Every occurrence of a term as Array(first word, last word).
Private Function TermHits(ByVal term As Variant, ByRef words() As String, ByVal index As Object, ByVal memo As Object) As Collection
    Dim key As String, first As String, hits As New Collection, candidates As New Collection, k As Variant, p As Variant, j As Long, ok As Boolean
    key = Join(term, " ")
    If memo.Exists(key) Then Set TermHits = memo(key): Exit Function
    first = CStr(term(0))
    If Right$(first, 1) = "*" Then
        For Each k In index.Keys
            If WordMatches(first, CStr(k)) Then
                For Each p In index(k): candidates.Add p: Next
            End If
        Next
    ElseIf index.Exists(first) Then
        For Each p In index(first): candidates.Add p: Next
    End If
    For Each p In candidates
        ok = (CLng(p) + UBound(term) <= UBound(words))
        j = 1
        Do While ok And j <= UBound(term)
            ok = WordMatches(CStr(term(j)), words(CLng(p) + j)): j = j + 1
        Loop
        If ok Then hits.Add Array(CLng(p), CLng(p) + UBound(term))
    Next
    memo.Add key, hits: Set TermHits = hits
End Function
'Words between two spans; 0 when they touch or overlap.
Private Function Gap(ByVal a As Variant, ByVal b As Variant) As Long
    If b(1) < a(0) Then
        Gap = a(0) - b(1)
    ElseIf b(0) > a(1) Then
        Gap = b(0) - a(1)
    End If
End Function
Private Function AnyNear(ByVal terms As Variant, ByVal span As Variant, ByVal reach As Long, ByRef words() As String, ByVal index As Object, ByVal memo As Object, ByRef found As Variant) As Boolean
    Dim i As Long, hit As Variant, term As Variant, beforeOnly As Boolean
    For i = 0 To UBound(terms)
        'A term starting with < counts only before the span ("screened 400 Sailors for
        'hypertension" is a duty; "on LIMDU, he trained" is not excused by "trained").
        term = terms(i): beforeOnly = (Left$(CStr(term(0)), 1) = "<")
        If beforeOnly Then term(0) = Mid$(CStr(term(0)), 2)
        For Each hit In TermHits(term, words, index, memo)
            If beforeOnly Then
                'A leading duty verb may sit a few words further back ("screened 400 Sailors for").
                If hit(1) < span(0) And Gap(span, hit) <= reach * 2 Then found = hit: AnyNear = True: Exit Function
            ElseIf Gap(span, hit) <= reach Then
                found = hit: AnyNear = True: Exit Function
            End If
        Next
    Next
End Function

'--- Exceptions and dismissals -----------------------------------------------------------
Private Sub LoadExceptions()
    Dim r As DAO.Recordset, starts() As Long, lengths() As Long, words() As String
    Set exceptionTerms = New Collection
    Set r = CurrentDb.OpenRecordset("SELECT Phrase,RuleID FROM RestrictedException WHERE RemovedAt Is Null", dbOpenSnapshot)
    Do While Not r.EOF
        words = Tokenize(Nz(r!Phrase, ""), starts, lengths)
        If Len(words(0)) > 0 Then exceptionTerms.Add Array(words, Nz(r!RuleID, ""))
        r.MoveNext
    Loop
    r.Close: exceptionsLoaded = True
End Sub
Private Sub LoadDismissals(ByVal reportID As Long)
    Dim r As DAO.Recordset
    Set dismissed = CreateObject("Scripting.Dictionary"): dismissedReport = reportID
    If reportID = 0 Then Exit Sub
    Set r = CurrentDb.OpenRecordset("SELECT RuleID,FieldKey,MatchText FROM RestrictedDismissal WHERE ReportID=" & reportID & " AND UndoneAt Is Null", dbOpenSnapshot)
    Do While Not r.EOF
        dismissed(r!RuleID & "|" & r!FieldKey & "|" & LCase$(r!MatchText)) = True: r.MoveNext
    Loop
    r.Close
End Sub
'Forget cached results after exceptions or dismissals change.
Public Sub RestrictedInvalidate()
    exceptionsLoaded = False: dismissedReport = -1
    Set results = Nothing: Set resultOrder = Nothing
End Sub
Private Function Excepted(ByVal ruleID As String, ByVal span As Variant, ByRef words() As String, ByVal index As Object, ByVal memo As Object) As Boolean
    Dim item As Variant, hit As Variant
    For Each item In exceptionTerms
        If item(1) = "" Or item(1) = ruleID Then
            For Each hit In TermHits(item(0), words, index, memo)
                If Gap(span, hit) = 0 Then Excepted = True: Exit Function
            Next
        End If
    Next
End Function

'--- Special kinds -----------------------------------------------------------------------
Private Function IsWordChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    IsWordChar = (c Like "[A-Za-z0-9]")
End Function
Private Sub AddRaw(ByVal hits As Collection, ByVal text As String, ByVal start As Long, ByVal length As Long)
    hits.Add Array(start, length, Mid$(text, start, length))
End Sub
'Classification markings, case-sensitive, not inside a word ("Sailor(S)" is not a marking).
Private Sub MarkingHits(ByVal terms As Variant, ByVal text As String, ByVal hits As Collection)
    Dim i As Long, p As Long, term As String, before As String, after As String
    For i = 0 To UBound(terms)
        term = CStr(terms(i)): p = InStr(1, text, term, vbBinaryCompare)
        Do While p > 0
            If p > 1 Then before = Mid$(text, p - 1, 1) Else before = ""
            after = Mid$(text, p + Len(term), 1)
            If Not IsWordChar(before) And (Not IsWordChar(after) Or Not IsWordChar(Right$(term, 1))) Then AddRaw hits, text, p, Len(term)
            p = InStr(p + 1, text, term, vbBinaryCompare)
        Loop
    Next
End Sub
'SSNs, DoD IDs and phone numbers (digit runs and dashed forms) and e-mail addresses.
Private Sub IdentifierHits(ByVal text As String, ByVal hits As Collection)
    Dim i As Long, j As Long, digits As Long, c As String, s As Long, e As Long, prior As String
    i = 1
    Do While i <= Len(text)
        c = Mid$(text, i, 1)
        prior = "": If i > 1 Then prior = Mid$(text, i - 1, 1)
        If c Like "#" And Not IsWordChar(prior) Then
            j = i: digits = 0
            Do While j <= Len(text)
                c = Mid$(text, j, 1)
                If c Like "#" Then
                    digits = digits + 1
                ElseIf (c = "-" Or c = "." Or c = " ") And j < Len(text) And Mid$(text, j + 1, 1) Like "#" And digits > 0 And digits < 10 Then
                    'Separators inside an identifier: 123-45-6789, 555.123.4567, 555 123 4567.
                Else
                    Exit Do
                End If
                j = j + 1
            Loop
            If (digits = 9 Or digits = 10) And Not IsWordChar(Mid$(text, j, 1)) Then AddRaw hits, text, i, j - i
            i = j
        ElseIf c = "@" Then
            s = i
            Do While s > 1
                If Not (IsWordChar(Mid$(text, s - 1, 1)) Or Mid$(text, s - 1, 1) Like "[._%+-]") Then Exit Do
                s = s - 1
            Loop
            e = i
            Do While e < Len(text)
                If Not (IsWordChar(Mid$(text, e + 1, 1)) Or Mid$(text, e + 1, 1) Like "[.-]") Then Exit Do
                e = e + 1
            Loop
            If s < i And InStr(i, Mid$(text, 1, e), ".") > i + 1 Then AddRaw hits, text, s, e - s + 1
            i = e + 1
        Else
            i = i + 1
        End If
    Loop
End Sub
'Four or more consecutive all-capital words with at least 20 capital letters.
Private Sub CapitalHits(ByVal text As String, ByRef words() As String, ByRef starts() As Long, ByRef lengths() As Long, ByVal hits As Collection)
    Dim i As Long, run As Long, letters As Long, first As Long, original As String, upper As Boolean, k As Long, c As String
    first = -1
    For i = 0 To UBound(words) + 1
        upper = False
        If i <= UBound(words) Then
            If Len(words(i)) > 0 Then
                original = Mid$(text, starts(i), lengths(i)): upper = (original = UCase$(original) And original <> LCase$(original))
            End If
        End If
        If upper Then
            If first < 0 Then first = i: run = 0: letters = 0
            run = run + 1
            For k = 1 To Len(original): c = Mid$(original, k, 1): If c Like "[A-Z]" Then letters = letters + 1
            Next
        Else
            If first >= 0 And run >= 4 And letters >= 20 Then AddRaw hits, text, starts(first), IIf(starts(i - 1) + lengths(i - 1) - starts(first) > 60, 60, starts(i - 1) + lengths(i - 1) - starts(first))
            first = -1
        End If
    Next
End Sub

'--- Matching ----------------------------------------------------------------------------
'Every match in one block as Array(rule ID, matched text, start, length), before
'dismissals. Exceptions are applied. Results are cached by kind, block and text.
Private Function BlockMatches(ByVal kind As String, ByVal key As String, ByVal text As String) As Collection
    Dim cacheKey As String, found As New Collection, seen As Object, words() As String, starts() As Long, lengths() As Long, index As Object, memo As Object
    Dim g As Long, row As Variant, t As Long, hit As Variant, near As Variant, ruleID As String, raw As New Collection, span As Variant, text0 As String, blocks As String, item As Variant
    Compile
    If Not exceptionsLoaded Then LoadExceptions: Set results = Nothing
    If results Is Nothing Then Set results = CreateObject("Scripting.Dictionary"): Set resultOrder = New Collection
    cacheKey = kind & "|" & key & "|" & text
    If results.Exists(cacheKey) Then Set BlockMatches = results(cacheKey): Exit Function
    Set seen = CreateObject("Scripting.Dictionary"): Set memo = CreateObject("Scripting.Dictionary")
    words = Tokenize(text, starts, lengths): Set index = WordIndex(words)
    For g = 0 To groupCount - 1
        row = ruleRows(groupRule(g)): ruleID = CStr(row(0)): blocks = CStr(row(8))
        If Len(blocks) = 0 Or InStr("|" & blocks & "|", "|" & key & "|") > 0 Then
            Select Case row(7)
                Case "marking"
                    Set raw = New Collection: MarkingHits groupTerms(g), text, raw
                    For Each item In raw: AddFound found, seen, ruleID, item(2), item(0), item(1): Next
                Case "caps"
                    Set raw = New Collection: CapitalHits text, words, starts, lengths, raw
                    For Each item In raw: AddFound found, seen, ruleID, item(2), item(0), item(1): Next
                Case Else
                    If row(7) = "pii" And g = FirstGroup(groupRule(g)) Then
                        Set raw = New Collection: IdentifierHits text, raw
                        For Each item In raw: AddFound found, seen, ruleID, item(2), item(0), item(1): Next
                    End If
                    For t = 0 To UBound(groupTerms(g))
                        For Each hit In TermHits(groupTerms(g)(t), words, index, memo)
                            span = hit
                            If UBound(groupNear(g)) >= 0 Then
                                If Not AnyNear(groupNear(g), hit, groupWithin(g), words, index, memo, near) Then GoTo NextHit
                                span = Array(IIf(near(0) < hit(0), near(0), hit(0)), IIf(near(1) > hit(1), near(1), hit(1)))
                            End If
                            If UBound(groupUnless(g)) >= 0 Then
                                If AnyNear(groupUnless(g), span, UNLESS_REACH, words, index, memo, near) Then GoTo NextHit
                            End If
                            If Excepted(ruleID, hit, words, index, memo) Then GoTo NextHit
                            text0 = Mid$(text, starts(hit(0)), starts(hit(1)) + lengths(hit(1)) - starts(hit(0)))
                            AddFound found, seen, ruleID, text0, starts(hit(0)), Len(text0)
NextHit:
                        Next
                    Next
            End Select
        End If
    Next
    If resultOrder.Count >= CACHE_LIMIT Then results.Remove resultOrder(1): resultOrder.Remove 1
    results.Add cacheKey, found: resultOrder.Add cacheKey
    Set BlockMatches = found
End Function
Private Function FirstGroup(ByVal ruleIndex As Long) As Long
    Dim g As Long
    For g = 0 To groupCount - 1
        If groupRule(g) = ruleIndex Then FirstGroup = g: Exit Function
    Next
End Function
'One finding per rule and matched words in a block.
Private Sub AddFound(ByVal found As Collection, ByVal seen As Object, ByVal ruleID As String, ByVal matched As String, ByVal start As Long, ByVal length As Long)
    Dim key As String
    key = ruleID & "|" & LCase$(matched)
    If seen.Exists(key) Then Exit Sub
    seen.Add key, True: found.Add Array(ruleID, matched, start, length)
End Sub
'Matches not dismissed for this report.
Public Function RestrictedMatches(ByVal kind As String, ByVal key As String, ByVal text As String, ByVal reportID As Long) As Collection
    Dim item As Variant, out As New Collection
    If dismissed Is Nothing Or dismissedReport <> reportID Then LoadDismissals reportID
    For Each item In BlockMatches(kind, key, text)
        If Not dismissed.Exists(item(0) & "|" & key & "|" & LCase$(item(1))) Then out.Add item
    Next
    Set RestrictedMatches = out
End Function
'Adds one Validation finding per match in the given narrative blocks.
Public Sub VRestrictedContent(ByVal findings As Collection, ByVal f As Object, ByVal id As Long, ByVal kind As String, ByVal keys As Collection)
    Dim key As Variant, text As String, item As Variant, row As Variant, message As String, finding As Object
    On Error GoTo Unavailable
    For Each key In keys
        If f.Exists(CStr(key)) Then
            text = Nz(f(CStr(key)), "")
            If Len(text) > 0 Then
                For Each item In RestrictedMatches(kind, CStr(key), text, id)
                    row = ruleRows(ruleIndexByID(CStr(item(0))))
                    message = "Possible " & row(1) & ": " & Chr$(34) & item(1) & Chr$(34) & ". " & _
                        IIf(row(6) = "Needs review", "Review before signing. ", "") & "Dismiss it for this report or allow the phrase if it is acceptable."
                    VAdd findings, f, id, CStr(key), CStr(row(6)), CStr(item(0)), message
                    Set finding = findings(findings.Count)
                    finding.Add "Match", item(1): finding.Add "Start", item(2): finding.Add "Length", item(3)
                Next
            End If
        End If
    Next
    Exit Sub
Unavailable:
    'A catalogue or workspace problem must not stop the other checks.
    VAdd findings, f, id, "OfficialComments", "Information", "APP-RESTRICTED", "The restricted-content check could not run: " & Err.Description
End Sub

'--- Dismissals and exceptions -----------------------------------------------------------
Public Function RestrictedDismiss(ByVal reportID As Long, ByVal ruleID As String, ByVal fieldKey As String, ByVal matched As String, ByVal reason As String) As Long
    Dim r As DAO.Recordset, op As Long
    Compile
    If reportID = 0 Then Err.Raise vbObjectError + 9851, "modRestrictedContent.RestrictedDismiss", "Save the report before dismissing a finding."
    'Not an enforced relationship: Report and UserProfile are at Access's 32-index limit.
    If PersistenceCount("Report", "ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 9851, "modRestrictedContent.RestrictedDismiss", "The report no longer exists."
    If Not ruleIndexByID.Exists(ruleID) Then Err.Raise vbObjectError + 9851, "modRestrictedContent.RestrictedDismiss", "Unknown restricted-content rule."
    If Len(matched) = 0 Or Len(matched) > 255 Then Err.Raise vbObjectError + 9851, "modRestrictedContent.RestrictedDismiss", "Choose a restricted-content finding."
    op = BeginOperation("Dismiss content finding", "Report " & reportID & ", " & ruleID & ", " & fieldKey)
    Set r = CurrentDb.OpenRecordset("RestrictedDismissal", dbOpenDynaset, dbAppendOnly)
    r.AddNew: r!ReportID = reportID: r!RuleID = ruleID: r!FieldKey = fieldKey: r!MatchText = matched
    r!Reason = Left$(Trim$(reason), 255): r!DismissedBy = MajorAuthor(): r!DismissedAt = Now()
    r.Update: r.Bookmark = r.LastModified: RestrictedDismiss = r!DismissalID: r.Close
    FinishOperation op, "Completed", "Finding dismissed for this report only; the report is unchanged."
    RestrictedInvalidate
End Function
Public Function RestrictedAllow(ByVal phrase As String, ByVal ruleID As String, ByVal reason As String) As Long
    Dim r As DAO.Recordset, op As Long, starts() As Long, lengths() As Long, words() As String
    Compile
    phrase = Trim$(phrase)
    words = Tokenize(phrase, starts, lengths)
    If Len(words(0)) = 0 Or Len(phrase) > 120 Then Err.Raise vbObjectError + 9852, "modRestrictedContent.RestrictedAllow", "Enter a phrase of up to 120 characters."
    If Len(ruleID) > 0 And Not ruleIndexByID.Exists(ruleID) Then Err.Raise vbObjectError + 9852, "modRestrictedContent.RestrictedAllow", "Unknown restricted-content rule."
    If Nz(MajorScalar("SELECT Count(*) FROM RestrictedException WHERE RemovedAt Is Null AND Phrase=" & QuoteSQL(phrase) & " AND Nz(RuleID,'')=" & QuoteSQL(ruleID)), 0) > 0 Then Err.Raise vbObjectError + 9852, "modRestrictedContent.RestrictedAllow", "That phrase is already allowed."
    op = BeginOperation("Allow content phrase", phrase & IIf(Len(ruleID) > 0, " (" & ruleID & ")", " (all rules)"))
    Set r = CurrentDb.OpenRecordset("RestrictedException", dbOpenDynaset, dbAppendOnly)
    r.AddNew: r!Phrase = phrase: If Len(ruleID) > 0 Then r!RuleID = ruleID
    r!Reason = Left$(Trim$(reason), 255): r!AddedBy = MajorAuthor(): r!AddedAt = Now()
    r.Update: r.Bookmark = r.LastModified: RestrictedAllow = r!ExceptionID: r.Close
    FinishOperation op, "Completed", "Phrase allowed for every report; no report changed."
    RestrictedInvalidate
End Function
'Each CurrentDb call returns a new object, so RecordsAffected is read from the same one.
Public Sub RestrictedRemoveException(ByVal exceptionID As Long)
    Dim db As DAO.Database, op As Long
    Set db = CurrentDb
    op = BeginOperation("Remove allowed content phrase", "Allowed phrase " & exceptionID)
    db.Execute "UPDATE RestrictedException SET RemovedAt=Now(),RemovedBy=" & MajorAuthor() & " WHERE ExceptionID=" & exceptionID & " AND RemovedAt Is Null", dbFailOnError
    If db.RecordsAffected = 0 Then FinishOperation op, "Failed", "Already removed.": Err.Raise vbObjectError + 9853, "modRestrictedContent.RestrictedRemoveException", "That allowed phrase was already removed."
    FinishOperation op, "Completed", "Phrase no longer allowed; its findings return."
    RestrictedInvalidate
End Sub
Public Sub RestrictedRestoreDismissal(ByVal dismissalID As Long)
    Dim db As DAO.Database, op As Long
    Set db = CurrentDb
    op = BeginOperation("Restore content finding", "Dismissal " & dismissalID)
    db.Execute "UPDATE RestrictedDismissal SET UndoneAt=Now(),UndoneBy=" & MajorAuthor() & " WHERE DismissalID=" & dismissalID & " AND UndoneAt Is Null", dbFailOnError
    If db.RecordsAffected = 0 Then FinishOperation op, "Failed", "Already restored.": Err.Raise vbObjectError + 9853, "modRestrictedContent.RestrictedRestoreDismissal", "That finding was already restored."
    FinishOperation op, "Completed", "Finding shown again on its report."
    RestrictedInvalidate
End Sub
'List rows for the allowed-phrases window: "E:id" or "D:id"; type; phrase or finding; detail.
Public Function RestrictedListRows() As String
    Dim r As DAO.Recordset, rows As String
    Set r = CurrentDb.OpenRecordset("SELECT e.ExceptionID,e.Phrase,e.RuleID,e.Reason,e.AddedAt,p.DisplayName FROM RestrictedException AS e LEFT JOIN UserProfile AS p ON e.AddedBy=p.UserProfileID WHERE e.RemovedAt Is Null ORDER BY e.Phrase", dbOpenSnapshot)
    Do While Not r.EOF
        rows = rows & IIf(rows = "", "", ";") & UXListCell("E:" & r!ExceptionID) & ";" & UXListCell("Allowed phrase") & ";" & UXListCell(r!Phrase) & ";" & UXListCell(IIf(IsNull(r!RuleID), "All rules", RestrictedRuleTitle(Nz(r!RuleID, ""))) & " - " & Nz(r!DisplayName, "") & ", " & Format$(r!AddedAt, "yyyy-mm-dd") & IIf(Nz(r!Reason, "") <> "", ": " & r!Reason, ""))
        r.MoveNext
    Loop
    r.Close
    Set r = CurrentDb.OpenRecordset("SELECT d.DismissalID,d.RuleID,d.MatchText,d.Reason,d.DismissedAt,p.DisplayName,r.OfficialName,r.MemberName FROM (RestrictedDismissal AS d LEFT JOIN UserProfile AS p ON d.DismissedBy=p.UserProfileID) LEFT JOIN Report AS r ON d.ReportID=r.ReportID WHERE d.UndoneAt Is Null ORDER BY d.DismissedAt DESC", dbOpenSnapshot)
    Do While Not r.EOF
        rows = rows & IIf(rows = "", "", ";") & UXListCell("D:" & r!DismissalID) & ";" & UXListCell("Dismissed: " & UXIdentity(Nz(r!OfficialName, ""), Nz(r!MemberName, ""))) & ";" & UXListCell(r!MatchText) & ";" & UXListCell(RestrictedRuleTitle(Nz(r!RuleID, "")) & " - " & Nz(r!DisplayName, "") & ", " & Format$(r!DismissedAt, "yyyy-mm-dd") & IIf(Nz(r!Reason, "") <> "", ": " & r!Reason, ""))
        r.MoveNext
    Loop
    r.Close
    RestrictedListRows = rows
End Function
