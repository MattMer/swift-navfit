Attribute VB_Name = "modPDFInflate"
Option Explicit
' Bounded RFC 1950/1951 decoder. No external executable, DLL package or script.
Private Type InflateTree
    left(0 To 640) As Integer
    right(0 To 640) As Integer
    symbol(0 To 640) As Integer
End Type
Private encoded() As Byte, decoded() As Byte, inputAt As Long, inputEnd As Long
Private bitBuffer As Long, bitCount As Long, outputAt As Long, outputLimit As Long
Private powers(0 To 24) As Long
Private Sub BadDeflate(ByVal reason As String)
    Err.Raise vbObjectError + 4750, "PDFInflateZlib", "Invalid or unsupported compressed PDF data: " & reason
End Sub
Private Function Bits(ByVal count As Long) As Long
    Do While bitCount < count
        If inputAt >= inputEnd Then BadDeflate "truncated stream"
        bitBuffer = bitBuffer Or (CLng(encoded(inputAt)) * powers(bitCount))
        inputAt = inputAt + 1: bitCount = bitCount + 8
    Loop
    Bits = bitBuffer And (powers(count) - 1)
    bitBuffer = bitBuffer \ powers(count): bitCount = bitCount - count
End Function
Private Sub MakeTree(ByRef tree As InflateTree, ByRef lengths() As Long, ByVal count As Long)
    Dim counts(0 To 15) As Long, nextCode(0 To 15) As Long
    Dim i As Long, width As Long, code As Long, node As Long, nodes As Long, child As Long, remaining As Long
    For i = 0 To 640: tree.left(i) = 0: tree.right(i) = 0: tree.symbol(i) = -1: Next
    For i = 0 To count - 1
        width = lengths(i)
        If width < 0 Or width > 15 Then BadDeflate "code length"
        If width > 0 Then counts(width) = counts(width) + 1
    Next
    remaining = 1
    For width = 1 To 15
        remaining = remaining * 2 - counts(width)
        If remaining < 0 Then BadDeflate "oversubscribed Huffman tree"
        code = (code + counts(width - 1)) * 2: nextCode(width) = code
    Next
    nodes = 1
    For i = 0 To count - 1
        width = lengths(i)
        If width > 0 Then
            code = nextCode(width): nextCode(width) = code + 1: node = 1
            For width = width - 1 To 0 Step -1
                If tree.symbol(node) >= 0 Then BadDeflate "Huffman prefix collision"
                If (code And powers(width)) = 0 Then child = tree.left(node) Else child = tree.right(node)
                If child = 0 Then
                    nodes = nodes + 1: If nodes > 640 Then BadDeflate "Huffman tree limit"
                    child = nodes
                    If (code And powers(width)) = 0 Then tree.left(node) = child Else tree.right(node) = child
                End If
                node = child
            Next
            If tree.symbol(node) >= 0 Or tree.left(node) <> 0 Or tree.right(node) <> 0 Then BadDeflate "duplicate Huffman code"
            tree.symbol(node) = i
        End If
    Next
End Sub
Private Function Symbol(ByRef tree As InflateTree) As Long
    Dim node As Long, depth As Long
    node = 1
    For depth = 1 To 16
        If tree.symbol(node) >= 0 Then Symbol = tree.symbol(node): Exit Function
        If Bits(1) = 0 Then node = tree.left(node) Else node = tree.right(node)
        If node = 0 Then BadDeflate "invalid Huffman code"
    Next
    BadDeflate "Huffman depth"
End Function
Private Sub PutByte(ByVal value As Long)
    If outputAt >= outputLimit Then BadDeflate "decoded size limit"
    decoded(outputAt) = value: outputAt = outputAt + 1
End Sub
Public Function PDFInflateZlib(ByRef source() As Byte, Optional ByVal limit As Long = 4000000) As Byte()
    Dim literalTree As InflateTree, distanceTree As InflateTree, codeTree As InflateTree
    Dim literalLengths(0 To 287) As Long, distanceLengths(0 To 31) As Long, codeLengths(0 To 18) As Long
    Dim allLengths(0 To 319) As Long, order As Variant, lengthBase As Variant, lengthExtra As Variant, distanceBase As Variant, distanceExtra As Variant
    Dim finalBlock As Long, kind As Long, count As Long, inverse As Long, i As Long, j As Long, value As Long
    Dim hlit As Long, hdist As Long, hclen As Long, total As Long, repeat As Long, previous As Long, length As Long, distance As Long
    Dim a As Long, b As Long, checksum As Double, failure As Long, message As String
    On Error GoTo Failed
    If limit < 1 Or limit > 4000000 Then BadDeflate "invalid output limit"
    encoded = source: inputEnd = UBound(encoded) - 3
    If LBound(encoded) <> 0 Or inputEnd < 2 Then BadDeflate "short zlib header"
    If (encoded(0) And 15) <> 8 Or encoded(0) > 120 Or ((CLng(encoded(0)) * 256 + encoded(1)) Mod 31) <> 0 Then BadDeflate "zlib header"
    If (encoded(1) And 32) <> 0 Then BadDeflate "preset dictionary"
    powers(0) = 1: For i = 1 To 24: powers(i) = powers(i - 1) * 2: Next
    inputAt = 2: bitBuffer = 0: bitCount = 0: outputAt = 0: outputLimit = limit: ReDim decoded(0 To limit - 1)
    order = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    lengthBase = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    lengthExtra = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    distanceBase = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577)
    distanceExtra = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13)
    Do
        finalBlock = Bits(1): kind = Bits(2)
        If kind = 0 Then
            bitBuffer = 0: bitCount = 0: count = Bits(16): inverse = Bits(16)
            If (count Xor 65535) <> inverse Then BadDeflate "stored block length"
            If count > outputLimit - outputAt Then BadDeflate "decoded size limit"
            For i = 1 To count: PutByte Bits(8): Next
        ElseIf kind = 1 Or kind = 2 Then
            Erase literalLengths: Erase distanceLengths
            If kind = 1 Then
                For i = 0 To 287
                    Select Case i
                        Case 0 To 143: literalLengths(i) = 8
                        Case 144 To 255: literalLengths(i) = 9
                        Case 256 To 279: literalLengths(i) = 7
                        Case Else: literalLengths(i) = 8
                    End Select
                Next
                For i = 0 To 31: distanceLengths(i) = 5: Next
            Else
                hlit = Bits(5) + 257: hdist = Bits(5) + 1: hclen = Bits(4) + 4
                If hlit > 286 Then BadDeflate "literal alphabet"
                Erase codeLengths
                For i = 0 To hclen - 1: codeLengths(order(i)) = Bits(3): Next
                MakeTree codeTree, codeLengths, 19
                total = hlit + hdist: i = 0
                Do While i < total
                    value = Symbol(codeTree)
                    If value < 16 Then
                        allLengths(i) = value: i = i + 1
                    Else
                        Select Case value
                            Case 16
                                If i = 0 Then BadDeflate "repeat without previous code"
                                previous = allLengths(i - 1): repeat = Bits(2) + 3
                            Case 17: previous = 0: repeat = Bits(3) + 3
                            Case 18: previous = 0: repeat = Bits(7) + 11
                            Case Else: BadDeflate "code length symbol"
                        End Select
                        If repeat > total - i Then BadDeflate "code length repeat"
                        For j = 1 To repeat: allLengths(i) = previous: i = i + 1: Next
                    End If
                Loop
                For i = 0 To hlit - 1: literalLengths(i) = allLengths(i): Next
                For i = 0 To hdist - 1: distanceLengths(i) = allLengths(hlit + i): Next
            End If
            If literalLengths(256) = 0 Then BadDeflate "missing end code"
            MakeTree literalTree, literalLengths, 288: MakeTree distanceTree, distanceLengths, 32
            Do
                value = Symbol(literalTree)
                If value < 256 Then
                    PutByte value
                ElseIf value = 256 Then
                    Exit Do
                Else
                    If value > 285 Then BadDeflate "reserved length symbol"
                    value = value - 257: length = lengthBase(value) + Bits(CLng(lengthExtra(value)))
                    value = Symbol(distanceTree): If value > 29 Then BadDeflate "reserved distance symbol"
                    distance = distanceBase(value) + Bits(CLng(distanceExtra(value)))
                    If distance > outputAt Or distance > 32768 Then BadDeflate "invalid backward distance"
                    If length > outputLimit - outputAt Then BadDeflate "decoded size limit"
                    For i = 1 To length: PutByte decoded(outputAt - distance): Next
                End If
            Loop
        Else
            BadDeflate "reserved block type"
        End If
    Loop Until finalBlock = 1
    If inputAt <> inputEnd Then BadDeflate "trailing compressed bytes"
    If outputAt = 0 Then BadDeflate "empty form packet"
    a = 1: b = 0
    For i = 0 To outputAt - 1
        a = a + decoded(i): If a >= 65521 Then a = a - 65521
        b = b + a: If b >= 65521 Then b = b - 65521
    Next
    checksum = CDbl(encoded(inputEnd)) * 16777216# + CDbl(encoded(inputEnd + 1)) * 65536# + CLng(encoded(inputEnd + 2)) * 256& + encoded(inputEnd + 3)
    If checksum <> CDbl(b) * 65536# + a Then BadDeflate "Adler-32 checksum"
    ReDim Preserve decoded(0 To outputAt - 1): PDFInflateZlib = decoded
    Erase encoded: Erase decoded: Exit Function
Failed:
    failure = Err.Number: message = Err.Description: Erase encoded: Erase decoded
    Err.Raise failure, "PDFInflateZlib", message
End Function
