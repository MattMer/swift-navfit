Attribute VB_Name = "modPDFLocal"
Option Explicit
' Read-only, bounded PDF 1.x object traversal. Never evaluates PDF/XFA actions.
' References: ISO 32000-1 sections 7.3-7.6 and 12.7.8; RFC 1950/1951.
Private fileBytes() As Byte, fileSize As Long, entries As Object, objects As Object, resolving As Object, objectStreams As Object
Private tokenCount As Long, decodedTotal As Long, encrypted As Boolean, fileKey() As Byte, streamCipher As String, stringCipher As String
Private Sub InvalidPDF(ByVal reason As String)
    Err.Raise vbObjectError + 4751, "PDFLocalPackets", reason
End Sub
Private Function Node(ByVal kind As String, Optional ByVal value As Variant) As Object
    Set Node = CreateObject("Scripting.Dictionary"): Node.Add "kind", kind
    If Not IsMissing(value) Then Node.Add "value", value
End Function
Private Function White(ByVal value As Long) As Boolean
    White = (value = 0 Or value = 9 Or value = 10 Or value = 12 Or value = 13 Or value = 32)
End Function
Private Function Delimiter(ByVal value As Long) As Boolean
    Delimiter = White(value) Or InStr(1, "()<>[]{}/%", ChrW$(value), vbBinaryCompare) > 0
End Function
Private Sub Space(ByRef data() As Byte, ByRef pos As Long)
    Do While pos <= UBound(data)
        If White(data(pos)) Then
            pos = pos + 1
        ElseIf data(pos) = 37 Then
            Do While pos <= UBound(data)
                If data(pos) = 10 Or data(pos) = 13 Then Exit Do
                pos = pos + 1
            Loop
        Else
            Exit Do
        End If
    Loop
End Sub
Private Function Word(ByRef data() As Byte, ByRef pos As Long) As String
    Dim start As Long, i As Long
    Space data, pos: start = pos
    Do While pos <= UBound(data)
        If Delimiter(data(pos)) Then Exit Do
        pos = pos + 1
        If pos - start > 128 Then InvalidPDF "PDF token length limit."
    Loop
    For i = start To pos - 1: Word = Word & ChrW$(data(i)): Next
End Function
Private Function UnsignedToken(ByVal text As String, ByVal limit As Long) As Boolean
    Dim i As Long, c As Long
    If Len(text) = 0 Or Len(text) > 10 Then Exit Function
    For i = 1 To Len(text): c = AscW(Mid$(text, i, 1)): If c < 48 Or c > 57 Then Exit Function
    Next
    UnsignedToken = (Val(text) <= limit)
End Function
Private Function HexDigit(ByVal value As Long) As Long
    Select Case value
        Case 48 To 57: HexDigit = value - 48
        Case 65 To 70: HexDigit = value - 55
        Case 97 To 102: HexDigit = value - 87
        Case Else: InvalidPDF "Invalid PDF hexadecimal digit."
    End Select
End Function
Private Function NameToken(ByRef data() As Byte, ByRef pos As Long) As String
    Dim c As Long
    If pos > UBound(data) Then InvalidPDF "Missing PDF name."
    If data(pos) <> 47 Then InvalidPDF "Expected PDF dictionary name."
    pos = pos + 1: NameToken = "/"
    Do While pos <= UBound(data)
        c = data(pos): If Delimiter(c) Then Exit Do
        pos = pos + 1
        If c = 35 Then
            If pos + 1 > UBound(data) Then InvalidPDF "Truncated escaped PDF name."
            c = HexDigit(data(pos)) * 16 + HexDigit(data(pos + 1)): pos = pos + 2
        End If
        If c = 0 Then InvalidPDF "Null in PDF name."
        NameToken = NameToken & ChrW$(c)
        If Len(NameToken) > 1024 Then InvalidPDF "PDF name length limit."
    Loop
End Function
Private Function StringToken(ByRef data() As Byte, ByRef pos As Long) As String
    Dim depth As Long, c As Long, value As Long, digits As Long, closed As Boolean, isHex As Boolean
    isHex = (data(pos) = 60): pos = pos + 1: depth = 1
    Do While pos <= UBound(data)
        c = data(pos): pos = pos + 1
        If isHex Then
            If c = 62 Then closed = True: Exit Do
            If Not White(c) Then
                value = HexDigit(c) * 16
                Do While pos <= UBound(data)
                    If Not White(data(pos)) Then Exit Do
                    pos = pos + 1
                Loop
                If pos > UBound(data) Then InvalidPDF "Truncated PDF hex string."
                If data(pos) <> 62 Then value = value + HexDigit(data(pos)): pos = pos + 1
                StringToken = StringToken & ChrW$(value)
            End If
        Else
            Select Case c
                Case 40: depth = depth + 1: StringToken = StringToken & "("
                Case 41
                    depth = depth - 1: If depth = 0 Then closed = True: Exit Do
                    StringToken = StringToken & ")"
                Case 92
                    If pos > UBound(data) Then InvalidPDF "Truncated PDF escape."
                    c = data(pos): pos = pos + 1
                    Select Case c
                        Case 110: c = 10
                        Case 114: c = 13
                        Case 116: c = 9
                        Case 98: c = 8
                        Case 102: c = 12
                        Case 10, 13
                            If c = 13 And pos <= UBound(data) Then
                                If data(pos) = 10 Then pos = pos + 1
                            End If
                            c = -1
                        Case 48 To 55
                            value = c - 48
                            For digits = 1 To 2
                                If pos > UBound(data) Then Exit For
                                If data(pos) < 48 Or data(pos) > 55 Then Exit For
                                value = value * 8 + data(pos) - 48: pos = pos + 1
                            Next
                            c = value Mod 256
                    End Select
                    If c >= 0 Then StringToken = StringToken & ChrW$(c)
                Case 13
                    If pos <= UBound(data) Then
                        If data(pos) = 10 Then pos = pos + 1
                    End If
                    StringToken = StringToken & vbLf
                Case Else: StringToken = StringToken & ChrW$(c)
            End Select
        End If
        If Len(StringToken) > 65536 Or depth > 64 Then InvalidPDF "PDF string limit."
    Loop
    If Not closed Then InvalidPDF "Unterminated PDF string."
End Function
Private Function ValueNode(ByRef data() As Byte, ByRef pos As Long, Optional ByVal depth As Long = 0, Optional ByVal owner As Long = 0, Optional ByVal generation As Long = 0) As Object
    Dim result As Object, item As Object, list As Collection, key As String, text As String, second As String, third As String
    Dim saved As Long, c As Long, i As Long, dots As Long, digits As Long
    tokenCount = tokenCount + 1
    If depth > 48 Or tokenCount > 500000 Then InvalidPDF "PDF structure limit."
    Space data, pos: If pos > UBound(data) Then InvalidPDF "Truncated PDF object."
    c = data(pos)
    Select Case c
        Case 47: Set result = Node("name", NameToken(data, pos))
        Case 40
            Set result = Node("string", StringToken(data, pos)): result.Add "owner", owner: result.Add "generation", generation
        Case 60
            If pos + 1 > UBound(data) Then InvalidPDF "Truncated PDF delimiter."
            If data(pos + 1) = 60 Then
                pos = pos + 2: Set result = Node("dict")
                Do
                    Space data, pos: If pos + 1 > UBound(data) Then InvalidPDF "Truncated PDF dictionary."
                    If data(pos) = 62 And data(pos + 1) = 62 Then pos = pos + 2: Exit Do
                    key = NameToken(data, pos)
                    If result.Exists(key) Then InvalidPDF "Duplicate PDF dictionary key."
                    Set item = ValueNode(data, pos, depth + 1, owner, generation): result.Add key, item
                Loop
            Else
                Set result = Node("string", StringToken(data, pos)): result.Add "owner", owner: result.Add "generation", generation
            End If
        Case 91
            pos = pos + 1: Set result = Node("array"): Set list = New Collection
            Do
                Space data, pos: If pos > UBound(data) Then InvalidPDF "Truncated PDF array."
                If data(pos) = 93 Then pos = pos + 1: Exit Do
                Set item = ValueNode(data, pos, depth + 1, owner, generation): list.Add item
                If list.Count > 100000 Then InvalidPDF "PDF array limit."
            Loop
            result.Add "items", list
        Case Else
            text = Word(data, pos)
            Select Case text
                Case "null": Set result = Node("null")
                Case "true", "false": Set result = Node("bool", (text = "true"))
                Case Else
                    For i = 1 To Len(text)
                        c = AscW(Mid$(text, i, 1))
                        If c >= 48 And c <= 57 Then
                            digits = digits + 1
                        ElseIf c = 46 Then
                            dots = dots + 1: If dots > 1 Then InvalidPDF "Invalid PDF number."
                        ElseIf i <> 1 Or (c <> 43 And c <> 45) Then
                            InvalidPDF "Invalid PDF value token."
                        End If
                    Next
                    If digits = 0 Or Abs(Val(text)) > 1000000000000# Then InvalidPDF "Invalid PDF number."
                    Set result = Node("number", Val(text))
                    If UnsignedToken(text, 200000) Then
                        saved = pos: second = Word(data, pos)
                        If UnsignedToken(second, 65535) Then
                            third = Word(data, pos)
                            If third = "R" Then
                                Set result = Node("ref", CLng(text)): result.Add "generation", CLng(second)
                            Else
                                pos = saved
                            End If
                        Else
                            pos = saved
                        End If
                    End If
            End Select
    End Select
    Set ValueNode = result
End Function
Private Function Number(ByVal item As Object, Optional ByVal limit As Long = 30000000, Optional ByVal negative As Boolean = False) As Long
    Dim value As Double
    If item Is Nothing Then InvalidPDF "Missing PDF number."
    If item("kind") <> "number" Then InvalidPDF "Expected direct PDF number."
    value = item("value")
    If value <> Fix(value) Or value > limit Then InvalidPDF "PDF integer limit."
    If negative Then
        If value < -2147483648# Then InvalidPDF "PDF integer limit."
    ElseIf value < 0 Then
        InvalidPDF "Negative PDF size or reference."
    End If
    Number = CLng(value)
End Function
Private Function Named(ByVal item As Object) As String
    If item("kind") <> "name" Then InvalidPDF "Expected PDF name."
    Named = item("value")
End Function
Private Function Field(ByVal item As Object, ByVal key As String) As Object
    If item Is Nothing Then InvalidPDF "Missing PDF dictionary."
    If Not item.Exists(key) Then InvalidPDF "Required PDF entry missing: " & key
    Set Field = item(key)
End Function
Private Function Indirect(ByVal offset As Long) As Object
    Dim pos As Long, numberText As String, generationText As String, result As Object, lengthNode As Object, length As Long
    If offset < 0 Or offset >= fileSize Then InvalidPDF "PDF object offset outside file."
    pos = offset: numberText = Word(fileBytes, pos): generationText = Word(fileBytes, pos)
    If Not UnsignedToken(numberText, 200000) Or Not UnsignedToken(generationText, 65535) Then InvalidPDF "Invalid PDF object header."
    If Word(fileBytes, pos) <> "obj" Then InvalidPDF "Missing PDF object header."
    Set result = ValueNode(fileBytes, pos, 0, CLng(numberText), CLng(generationText))
    result.Add "object", CLng(numberText): result.Add "objectGeneration", CLng(generationText)
    Space fileBytes, pos
    If result("kind") = "dict" Then
        length = pos
        If Word(fileBytes, pos) = "stream" Then
            If pos >= fileSize Then InvalidPDF "Truncated PDF stream."
            If fileBytes(pos) = 13 Then
                pos = pos + 1
                If pos < fileSize Then
                    If fileBytes(pos) = 10 Then pos = pos + 1
                End If
            ElseIf fileBytes(pos) = 10 Then
                pos = pos + 1
            Else
                InvalidPDF "Missing PDF stream line ending."
            End If
            Set lengthNode = Field(result, "/Length")
            If lengthNode("kind") = "ref" Then Set lengthNode = Resolve(lengthNode)
            length = Number(lengthNode)
            If length > fileSize - pos Then InvalidPDF "PDF stream exceeds file."
            result("kind") = "stream": result.Add "start", pos: result.Add "length", length
            pos = pos + length
            If Word(fileBytes, pos) <> "endstream" Then InvalidPDF "PDF stream length does not match terminator."
        Else
            pos = length
        End If
    End If
    If Word(fileBytes, pos) <> "endobj" Then InvalidPDF "Missing PDF object terminator."
    Set Indirect = result
End Function
Private Function Slice(ByRef source() As Byte, ByVal start As Long, ByVal count As Long) As Byte()
    Dim result() As Byte, i As Long
    If count <= 0 Or start < 0 Then InvalidPDF "Empty or invalid PDF byte range."
    If count > UBound(source) + 1 - start Then InvalidPDF "PDF byte range outside buffer."
    ReDim result(0 To count - 1)
    For i = 0 To count - 1: result(i) = source(start + i): Next
    Slice = result
End Function
Private Function Predictor(ByRef raw() As Byte, ByVal parameters As Object) As Byte()
    Dim method As Long, columns As Long, colors As Long, bits As Long, row As Long, width As Long, pixel As Long, i As Long, inputAt As Long, outputAt As Long
    Dim filter As Long, left As Long, up As Long, upperLeft As Long, value As Long, estimate As Long, pa As Long, pb As Long, pc As Long, result() As Byte
    method = 1: columns = 1: colors = 1: bits = 8
    If parameters.Exists("/Predictor") Then method = Number(parameters("/Predictor"), 15)
    If method = 1 Then Predictor = raw: Exit Function
    If parameters.Exists("/Columns") Then columns = Number(parameters("/Columns"), 100000)
    If parameters.Exists("/Colors") Then colors = Number(parameters("/Colors"), 32)
    If parameters.Exists("/BitsPerComponent") Then bits = Number(parameters("/BitsPerComponent"), 16)
    If bits <> 8 Or columns = 0 Or colors = 0 Then InvalidPDF "Unsupported PDF predictor geometry."
    width = columns * colors: pixel = colors
    If method = 2 Then
        If (UBound(raw) + 1) Mod width <> 0 Then InvalidPDF "Truncated TIFF predictor."
        result = raw
        For i = 0 To UBound(result)
            If i Mod width >= pixel Then result(i) = (CLng(result(i)) + result(i - pixel)) Mod 256
        Next
        Predictor = result: Exit Function
    End If
    If method < 10 Or method > 15 Then InvalidPDF "Unsupported PDF predictor."
    If (UBound(raw) + 1) Mod (width + 1) <> 0 Then InvalidPDF "Truncated PNG predictor."
    ReDim result(0 To ((UBound(raw) + 1) \ (width + 1)) * width - 1)
    Do While inputAt <= UBound(raw)
        filter = raw(inputAt): inputAt = inputAt + 1
        If filter > 4 Then InvalidPDF "Invalid PNG predictor filter."
        For i = 0 To width - 1
            left = 0: up = 0: upperLeft = 0
            If i >= pixel Then left = result(outputAt - pixel)
            If row > 0 Then up = result(outputAt - width)
            If row > 0 And i >= pixel Then upperLeft = result(outputAt - width - pixel)
            value = raw(inputAt): inputAt = inputAt + 1
            Select Case filter
                Case 1: value = value + left
                Case 2: value = value + up
                Case 3: value = value + (left + up) \ 2
                Case 4
                    estimate = left + up - upperLeft: pa = Abs(estimate - left): pb = Abs(estimate - up): pc = Abs(estimate - upperLeft)
                    If pa <= pb And pa <= pc Then
                        value = value + left
                    ElseIf pb <= pc Then
                        value = value + up
                    Else
                        value = value + upperLeft
                    End If
            End Select
            result(outputAt) = value Mod 256: outputAt = outputAt + 1
        Next
        row = row + 1
    Loop
    Predictor = result
End Function
Private Function StreamBytes(ByVal item As Object) As Byte()
    Dim result() As Byte, filter As Object, parameters As Object, list As Collection
    If item("kind") <> "stream" Then InvalidPDF "XFA data must be a PDF stream."
    If item.Exists("/F") Then InvalidPDF "External PDF streams are not supported."
    result = Slice(fileBytes, item("start"), item("length"))
    If encrypted Then
        If Not item.Exists("/Type") Then
            result = PDFDecryptObject(result, fileKey, item("object"), item("objectGeneration"), streamCipher)
        ElseIf Named(item("/Type")) <> "/XRef" Then
            result = PDFDecryptObject(result, fileKey, item("object"), item("objectGeneration"), streamCipher)
        End If
    End If
    If item.Exists("/Filter") Then
        Set filter = Resolve(item("/Filter"))
        If filter("kind") = "array" Then
            Set list = filter("items"): If list.Count <> 1 Then InvalidPDF "Unsupported PDF stream filter chain."
            Set filter = Resolve(list(1))
        End If
        If Named(filter) <> "/FlateDecode" And Named(filter) <> "/Fl" Then InvalidPDF "Unsupported PDF stream compression."
        result = PDFInflateZlib(result)
        If item.Exists("/DecodeParms") Then
            Set parameters = Resolve(item("/DecodeParms"))
            If parameters("kind") = "array" Then
                Set list = parameters("items"): If list.Count <> 1 Then InvalidPDF "Unsupported PDF filter parameters."
                Set parameters = Resolve(list(1))
            End If
            If parameters("kind") <> "null" Then
                If parameters("kind") <> "dict" Then InvalidPDF "Invalid PDF filter parameters."
                result = Predictor(result, parameters)
            End If
        End If
    End If
    If UBound(result) >= 4000000 Then InvalidPDF "Decoded PDF packet exceeds size limit."
    decodedTotal = decodedTotal + UBound(result) + 1
    If decodedTotal > 32000000 Then InvalidPDF "Aggregate PDF decode limit."
    StreamBytes = result
End Function
Private Function Resolve(ByVal item As Object) As Object
    Dim id As Long, key As String, entry As Variant, result As Object, stream As Object, data() As Byte, list As Object
    Dim n As Long, first As Long, pos As Long, i As Long, objectNumber As String, relative As String, target As Long, previous As Long
    If item("kind") <> "ref" Then Set Resolve = item: Exit Function
    id = item("value"): key = CStr(id)
    If Not entries.Exists(key) Then InvalidPDF "Missing active PDF object reference."
    entry = Split(entries(key), "|")
    If entry(0) = "0" Then InvalidPDF "Reference to a freed PDF object."
    If entry(0) = "1" And CLng(entry(2)) <> item("generation") Then InvalidPDF "Stale PDF object generation."
    If entry(0) = "2" And item("generation") <> 0 Then InvalidPDF "Invalid compressed object generation."
    If objects.Exists(key) Then Set Resolve = objects(key): Exit Function
    If resolving.Exists(key) Or resolving.Count > 64 Then InvalidPDF "Cyclic PDF object references."
    resolving.Add key, True
    If entry(0) = "1" Then
        Set result = Indirect(CLng(entry(1)))
        If result("object") <> id Or result("objectGeneration") <> item("generation") Then InvalidPDF "PDF xref/object mismatch."
    ElseIf entry(0) = "2" Then
        Set stream = Node("ref", CLng(entry(1))): stream.Add "generation", 0: Set stream = Resolve(stream)
        If Named(Field(stream, "/Type")) <> "/ObjStm" Then InvalidPDF "Invalid PDF object-stream reference."
        If Not objectStreams.Exists(entry(1)) Then
            data = StreamBytes(stream): objectStreams.Add entry(1), data
        Else
            data = objectStreams(entry(1))
        End If
        n = Number(Field(stream, "/N"), 100000): first = Number(Field(stream, "/First"), UBound(data) + 1)
        If n = 0 Or CLng(entry(2)) >= n Then InvalidPDF "Invalid object-stream index."
        pos = 0: target = -1: previous = -1
        For i = 0 To n - 1
            objectNumber = Word(data, pos): relative = Word(data, pos)
            If Not UnsignedToken(objectNumber, 200000) Or Not UnsignedToken(relative, UBound(data)) Then InvalidPDF "Invalid object-stream header."
            If CLng(relative) <= previous Then InvalidPDF "Unordered object-stream offsets."
            previous = CLng(relative)
            If i = CLng(entry(2)) Then
                If CLng(objectNumber) <> id Then InvalidPDF "Object-stream ID mismatch."
                target = first + CLng(relative)
            End If
        Next
        Space data, pos: If pos <> first Or target < first Or target > UBound(data) Then InvalidPDF "Invalid object-stream body."
        Set result = ValueNode(data, target)
    Else
        InvalidPDF "Unsupported PDF xref entry."
    End If
    resolving.Remove key: objects.Add key, result: Set Resolve = result
End Function
Private Function BigNumber(ByRef data() As Byte, ByRef pos As Long, ByVal width As Long) As Long
    Dim value As Double, i As Long
    If width < 0 Or width > 4 Or pos + width > UBound(data) + 1 Then InvalidPDF "Invalid PDF xref field."
    For i = 1 To width: value = value * 256# + data(pos): pos = pos + 1: Next
    If value > 2147483647# Then InvalidPDF "PDF xref integer limit."
    BigNumber = CLng(value)
End Function
Private Function XrefStream(ByVal offset As Long, ByVal sectionEntries As Object) As Object
    Dim stream As Object, widths As Collection, spans As Collection, data() As Byte, pos As Long, i As Long, j As Long, start As Long, count As Long
    Dim w0 As Long, w1 As Long, w2 As Long, kind As Long, a As Long, b As Long, key As String
    Set stream = Indirect(offset)
    If Named(Field(stream, "/Type")) <> "/XRef" Then InvalidPDF "Expected cross-reference stream."
    Set widths = Field(stream, "/W")("items")
    If widths.Count <> 3 Then InvalidPDF "Invalid xref stream widths."
    w0 = Number(widths(1), 4): w1 = Number(widths(2), 4): w2 = Number(widths(3), 4)
    If w0 + w1 + w2 = 0 Then InvalidPDF "Empty xref stream record."
    If stream.Exists("/Index") Then
        Set spans = stream("/Index")("items")
    Else
        Set spans = New Collection: spans.Add Node("number", 0): spans.Add Field(stream, "/Size")
    End If
    If spans.Count Mod 2 <> 0 Or spans.Count = 0 Then InvalidPDF "Invalid xref stream ranges."
    data = StreamBytes(stream)
    For i = 1 To spans.Count Step 2
        start = Number(spans(i), 200000): count = Number(spans(i + 1), 200000)
        If start + count > 200000 Then InvalidPDF "PDF object count limit."
        For j = start To start + count - 1
            kind = BigNumber(data, pos, w0): If w0 = 0 Then kind = 1
            a = BigNumber(data, pos, w1): b = BigNumber(data, pos, w2): key = CStr(j)
            If kind > 2 Then InvalidPDF "Unsupported xref entry type."
            If sectionEntries.Exists(key) Then sectionEntries.Remove key
            sectionEntries.Add key, CStr(kind) & "|" & a & "|" & b
        Next
    Next
    If pos <> UBound(data) + 1 Then InvalidPDF "Xref stream size mismatch."
    Set XrefStream = stream
End Function
Private Function CrossReferences(ByVal offset As Long) As Object
    Dim pos As Long, start As String, count As String, a As String, b As String, flag As String, i As Long, key As Variant
    Dim trailer As Object, extra As Object, sectionEntries As Object, visited As Object, root As Object, security As Object, documentID As Object
    Set visited = CreateObject("Scripting.Dictionary")
    Do
        If offset < 0 Or offset >= fileSize Or visited.Count >= 128 Then InvalidPDF "Invalid PDF revision chain."
        If visited.Exists(CStr(offset)) Then InvalidPDF "Cyclic PDF revision chain."
        visited.Add CStr(offset), True: Set sectionEntries = CreateObject("Scripting.Dictionary")
        pos = offset
        If Word(fileBytes, pos) = "xref" Then
            Do
                start = Word(fileBytes, pos)
                If start = "trailer" Then Exit Do
                count = Word(fileBytes, pos)
                If Not UnsignedToken(start, 200000) Or Not UnsignedToken(count, 200000) Then InvalidPDF "Invalid PDF xref table."
                If CDbl(start) + CDbl(count) > 200000 Then InvalidPDF "PDF object count limit."
                For i = CLng(start) To CLng(start) + CLng(count) - 1
                    a = Word(fileBytes, pos): b = Word(fileBytes, pos): flag = Word(fileBytes, pos)
                    If Not UnsignedToken(a, 30000000) Or Not UnsignedToken(b, 65535) Then InvalidPDF "Invalid PDF xref offset."
                    If flag <> "n" And flag <> "f" Then InvalidPDF "Invalid PDF xref state."
                    If sectionEntries.Exists(CStr(i)) Then InvalidPDF "Overlapping PDF xref ranges."
                    sectionEntries.Add CStr(i), IIf(flag = "n", "1", "0") & "|" & CLng(a) & "|" & CLng(b)
                Next
            Loop
            Set trailer = ValueNode(fileBytes, pos)
            If trailer("kind") <> "dict" Then InvalidPDF "Invalid PDF trailer."
            If trailer.Exists("/XRefStm") Then
                Set extra = XrefStream(Number(trailer("/XRefStm")), sectionEntries)
            End If
        Else
            Set trailer = XrefStream(offset, sectionEntries)
        End If
        For Each key In sectionEntries.Keys
            If Not entries.Exists(key) Then entries.Add CStr(key), sectionEntries(key)
        Next
        If root Is Nothing And trailer.Exists("/Root") Then Set root = trailer("/Root")
        If security Is Nothing And trailer.Exists("/Encrypt") Then Set security = trailer("/Encrypt")
        If documentID Is Nothing And trailer.Exists("/ID") Then Set documentID = trailer("/ID")
        If Not trailer.Exists("/Prev") Then Exit Do
        offset = Number(trailer("/Prev"))
    Loop
    If root Is Nothing Then InvalidPDF "No active PDF document catalog."
    Set CrossReferences = Node("trailer"): CrossReferences.Add "/Root", root
    If Not security Is Nothing Then CrossReferences.Add "/Encrypt", security
    If Not documentID Is Nothing Then CrossReferences.Add "/ID", documentID
End Function
Private Function Cipher(ByVal security As Object, ByVal key As String) As String
    Dim name As String, filters As Object, filter As Object
    name = "/Identity": If security.Exists(key) Then name = Named(security(key))
    If name = "/Identity" Then Cipher = "Identity": Exit Function
    Set filters = Field(security, "/CF"): Set filter = Field(filters, name)
    name = Named(Field(filter, "/CFM"))
    If name <> "/V2" And name <> "/AESV2" Then InvalidPDF "Unsupported PDF security filter."
    Cipher = Mid$(name, 2)
End Function
Private Sub InitializeSecurity(ByVal trailer As Object)
    Dim security As Object, id As Object, list As Collection, r As Long, version As Long, size As Long, metadata As Boolean
    If Not trailer.Exists("/Encrypt") Then Exit Sub
    Set security = Resolve(trailer("/Encrypt"))
    If security("kind") = "null" Then Exit Sub
    If Named(Field(security, "/Filter")) <> "/Standard" Then InvalidPDF "Unsupported PDF security handler."
    r = Number(Field(security, "/R"), 6): version = Number(Field(security, "/V"), 5)
    If r < 2 Or r > 4 Or (version <> 1 And version <> 2 And version <> 4) Then InvalidPDF "Unsupported PDF encryption; local import supports extraction-permitted R2-R4 forms that open without a password."
    size = 40: If security.Exists("/Length") Then size = Number(security("/Length"), 128)
    metadata = True: If security.Exists("/EncryptMetadata") Then metadata = security("/EncryptMetadata")("value")
    Set id = Field(trailer, "/ID"): Set list = id("items")
    If list.Count <> 2 Then InvalidPDF "Missing PDF encryption document ID."
    fileKey = PDFEmptyPasswordKey(r, size, Number(Field(security, "/P"), 2147483647, True), Field(security, "/O")("value"), Field(security, "/U")("value"), list(1)("value"), metadata)
    streamCipher = "V2": stringCipher = "V2"
    If version = 4 Then streamCipher = Cipher(security, "/StmF"): stringCipher = Cipher(security, "/StrF")
    encrypted = True
End Sub
Private Function PacketName(ByVal item As Object) As String
    Dim data() As Byte
    Set item = Resolve(item)
    If item("kind") <> "string" Then InvalidPDF "Invalid XFA packet name."
    PacketName = item("value")
    If encrypted And item("owner") > 0 Then
        data = PDFBinaryBytes(PacketName): data = PDFDecryptObject(data, fileKey, item("owner"), item("generation"), stringCipher): PacketName = PDFByteString(data)
    End If
End Function
Public Function PDFLocalPackets(ByVal path As String) As Object
    Dim file As Integer, opened As Boolean, tail As String, pos As Long, marker As Long, offset As String, trailer As Object, root As Object, form As Object, xfa As Object
    Dim list As Collection, names As Object, result As Object, item As Object, i As Long, key As String, raw() As Byte, before As String, failure As Long, message As String
    On Error GoTo Failed
    file = FreeFile: Open path For Binary Access Read Lock Write As #file: opened = True
    fileSize = LOF(file)
    If fileSize < 16 Or fileSize > 30000000 Then InvalidPDF "PDF is empty or exceeds the 30 MB import limit."
    ReDim fileBytes(0 To fileSize - 1): Get #file, , fileBytes
    before = PDFBytesHash(fileBytes)
    If fileBytes(0) <> 37 Or fileBytes(1) <> 80 Or fileBytes(2) <> 68 Or fileBytes(3) <> 70 Or fileBytes(4) <> 45 Then InvalidPDF "Not a supported PDF file."
    pos = fileSize - 4096: If pos < 0 Then pos = 0
    For i = pos To fileSize - 1: tail = tail & ChrW$(fileBytes(i)): Next
    marker = InStrRev(tail, "startxref", -1, vbBinaryCompare)
    If marker = 0 Then InvalidPDF "Missing final PDF cross-reference pointer."
    pos = pos + marker - 1 + 9: offset = Word(fileBytes, pos)
    If Not UnsignedToken(offset, fileSize - 1) Then InvalidPDF "Invalid final PDF cross-reference pointer."
    Do While pos < fileSize
        If Not White(fileBytes(pos)) Then Exit Do
        pos = pos + 1
    Loop
    If pos + 5 > fileSize Then InvalidPDF "Missing final PDF end marker."
    For i = 0 To 4
        If fileBytes(pos + i) <> AscW(Mid$("%%EOF", i + 1, 1)) Then InvalidPDF "Invalid final PDF end marker."
    Next
    pos = pos + 5
    Do While pos < fileSize
        If Not White(fileBytes(pos)) Then InvalidPDF "Unexpected data after final PDF revision."
        pos = pos + 1
    Loop
    tokenCount = 0: decodedTotal = 0: encrypted = False
    Set entries = CreateObject("Scripting.Dictionary"): Set objects = CreateObject("Scripting.Dictionary")
    Set resolving = CreateObject("Scripting.Dictionary"): Set objectStreams = CreateObject("Scripting.Dictionary")
    Set trailer = CrossReferences(CLng(offset)): InitializeSecurity trailer
    Set root = Resolve(Field(trailer, "/Root"))
    If Named(Field(root, "/Type")) <> "/Catalog" Then InvalidPDF "Invalid PDF document catalog."
    Set form = Resolve(Field(root, "/AcroForm"))
    If Not form.Exists("/XFA") Then InvalidPDF "No supported XFA form data. Flattened/scanned PDFs require manual entry."
    Set xfa = Resolve(form("/XFA"))
    If xfa("kind") <> "array" Then InvalidPDF "Unsupported XFA layout: expected separate template and data packets."
    Set list = xfa("items")
    If list.Count = 0 Or list.Count Mod 2 <> 0 Or list.Count > 128 Then InvalidPDF "Invalid XFA packet array."
    Set names = CreateObject("Scripting.Dictionary"): Set result = CreateObject("Scripting.Dictionary")
    For i = 1 To list.Count Step 2
        key = PacketName(list(i))
        If names.Exists(key) Then InvalidPDF "Ambiguous duplicate XFA packet."
        names.Add key, True
        Set item = Resolve(list(i + 1))
        If item("kind") <> "stream" Then InvalidPDF "Invalid XFA packet reference."
        If key = "template" Or key = "datasets" Then
            raw = StreamBytes(item): result.Add key, raw
        End If
    Next
    If Not result.Exists("template") Or Not result.Exists("datasets") Then InvalidPDF "Required XFA template/data packet missing."
    Close #file: opened = False
    If SourceFileHash(path) <> before Then InvalidPDF "PDF changed during local extraction."
    result.Add "sourceHash", before: Set PDFLocalPackets = result
    ClearLocalPDF: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If opened Then Close #file
    ClearLocalPDF
    On Error GoTo 0
    Err.Raise failure, "PDFLocalPackets", message
End Function
Private Sub ClearLocalPDF()
    Erase fileBytes: Erase fileKey
    Set entries = Nothing: Set objects = Nothing: Set resolving = Nothing: Set objectStreams = Nothing
    encrypted = False
End Sub
Public Function PDFPacketText(ByRef data() As Byte) As String
    Dim pos As Long, c As Long, extra As Long, minimum As Long, j As Long, value As Long, text As String, at As Long
    text = String$(UBound(data) + 1, vbNullChar): pos = 0
    If UBound(data) >= 2 Then
        If data(0) = 239 And data(1) = 187 And data(2) = 191 Then pos = 3
    End If
    Do While pos <= UBound(data)
        c = data(pos): pos = pos + 1
        Select Case c
            Case 0 To 127: extra = 0: minimum = 0
            Case 194 To 223: extra = 1: minimum = 128: c = c And 31
            Case 224 To 239: extra = 2: minimum = 2048: c = c And 15
            Case 240 To 244: extra = 3: minimum = 65536: c = c And 7
            Case Else: InvalidPDF "Invalid UTF-8 in XFA data."
        End Select
        For j = 1 To extra
            If pos > UBound(data) Then InvalidPDF "Truncated UTF-8 in XFA data."
            value = data(pos): pos = pos + 1
            If value < 128 Or value > 191 Then InvalidPDF "Invalid UTF-8 continuation."
            c = c * 64 + (value And 63)
        Next
        If c < minimum Or c > 1114111 Or (c >= 55296 And c <= 57343) Then InvalidPDF "Invalid XFA Unicode scalar."
        If c > 65535 Then
            c = c - 65536: at = at + 1: Mid$(text, at, 1) = ChrW$(55296 + c \ 1024)
            at = at + 1: Mid$(text, at, 1) = ChrW$(56320 + c Mod 1024)
        Else
            at = at + 1: Mid$(text, at, 1) = ChrW$(c)
        End If
    Loop
    PDFPacketText = Left$(text, at)
End Function
