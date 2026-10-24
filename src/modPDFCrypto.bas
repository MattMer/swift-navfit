Attribute VB_Name = "modPDFCrypto"
Option Explicit
' Windows cryptography only. PDF Standard Security, empty user password, R2-R4.
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal parameter As Long, ByRef data As Byte, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32" (ByVal provider As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal publicKey As LongPtr, ByVal flags As Long, ByRef key As LongPtr) As Long
Private Declare PtrSafe Function CryptSetKeyParam Lib "advapi32" (ByVal key As LongPtr, ByVal parameter As Long, ByRef data As Any, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDecrypt Lib "advapi32" (ByVal key As LongPtr, ByVal hash As LongPtr, ByVal final As Long, ByVal flags As Long, ByRef data As Byte, ByRef length As Long) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32" (ByVal key As LongPtr) As Long
Private Function Digest(ByRef data() As Byte, ByVal count As Long, ByVal algorithm As Long, ByVal size As Long) As Byte()
    Dim provider As LongPtr, hash As LongPtr, result() As Byte, failure As Long, message As String
    On Error GoTo Failed
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5, "modPDFCrypto.Digest", "Windows hashing unavailable."
    If CryptCreateHash(provider, algorithm, 0, 0, hash) = 0 Then Err.Raise 5, "modPDFCrypto.Digest", "Cannot initialize PDF hashing."
    If count > 0 Then
        If CryptHashData(hash, data(0), count, 0) = 0 Then Err.Raise 5, "modPDFCrypto.Digest", "Cannot hash PDF bytes."
    End If
    ReDim result(0 To size - 1)
    If CryptGetHashParam(hash, 2, result(0), size, 0) = 0 Then Err.Raise 5, "modPDFCrypto.Digest", "Cannot read PDF hash."
    Digest = result
    CryptDestroyHash hash: CryptReleaseContext provider, 0: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    If hash <> 0 Then CryptDestroyHash hash
    If provider <> 0 Then CryptReleaseContext provider, 0
    Err.Raise failure, "PDFDigest", message
End Function
Public Function PDFBytesHash(ByRef data() As Byte) As String
    Dim result() As Byte, i As Long
    result = Digest(data, UBound(data) + 1, &H800C&, 32)
    For i = 0 To 31: PDFBytesHash = PDFBytesHash & Right$("0" & Hex$(result(i)), 2): Next
End Function
Private Function MD5(ByRef data() As Byte, ByVal count As Long) As Byte()
    MD5 = Digest(data, count, &H8003&, 16)
End Function
Public Function PDFBinaryBytes(ByVal text As String) As Byte()
    Dim data() As Byte, i As Long
    If Len(text) = 0 Then Err.Raise 5, "modPDFCrypto.PDFBinaryBytes", "Missing PDF byte string."
    ReDim data(0 To Len(text) - 1)
    For i = 1 To Len(text): data(i - 1) = AscW(Mid$(text, i, 1)): Next
    PDFBinaryBytes = data
End Function
Public Function PDFByteString(ByRef data() As Byte) As String
    Dim i As Long, text As String
    text = String$(UBound(data) + 1, vbNullChar)
    For i = 0 To UBound(data): Mid$(text, i + 1, 1) = ChrW$(data(i)): Next
    PDFByteString = text
End Function
Private Function RC4(ByRef data() As Byte, ByRef key() As Byte) As Byte()
    Dim state(0 To 255) As Long, i As Long, j As Long, k As Long, swap As Long, result() As Byte
    For i = 0 To 255: state(i) = i: Next
    For i = 0 To 255
        j = (j + state(i) + key(i Mod (UBound(key) + 1))) Mod 256
        swap = state(i): state(i) = state(j): state(j) = swap
    Next
    ReDim result(0 To UBound(data)): i = 0: j = 0
    For k = 0 To UBound(data)
        i = (i + 1) Mod 256: j = (j + state(i)) Mod 256
        swap = state(i): state(i) = state(j): state(j) = swap
        result(k) = data(k) Xor state((state(i) + state(j)) Mod 256)
    Next
    RC4 = result
End Function
Public Function PDFEmptyPasswordKey(ByVal revision As Long, ByVal bits As Long, ByVal permissions As Long, ByVal owner As String, ByVal user As String, ByVal documentID As String, ByVal metadata As Boolean) As Byte()
    Dim pad As Variant, buffer() As Byte, digestBytes() As Byte, key() As Byte, probe() As Byte, roundKey() As Byte
    Dim i As Long, j As Long, count As Long, size As Long, unsigned As Double
    If revision < 2 Or revision > 4 Or bits < 40 Or bits > 128 Or bits Mod 8 <> 0 Then Err.Raise 5, "modPDFCrypto.PDFEmptyPasswordKey", "Unsupported PDF encryption revision."
    If (permissions And 16) = 0 Then Err.Raise 5, "modPDFCrypto.PDFEmptyPasswordKey", "This PDF does not permit data extraction."
    If Len(owner) <> 32 Or Len(user) <> 32 Or Len(documentID) = 0 Then Err.Raise 5, "modPDFCrypto.PDFEmptyPasswordKey", "Invalid PDF security dictionary."
    pad = Array(40, 191, 78, 94, 78, 117, 138, 65, 100, 0, 78, 86, 255, 250, 1, 8, 46, 46, 0, 182, 208, 104, 62, 128, 47, 12, 169, 254, 100, 83, 105, 122)
    count = 68 + Len(documentID): If revision >= 4 And Not metadata Then count = count + 4
    ReDim buffer(0 To count - 1)
    For i = 0 To 31: buffer(i) = pad(i): buffer(32 + i) = AscW(Mid$(owner, i + 1, 1)): Next
    unsigned = permissions: If unsigned < 0 Then unsigned = unsigned + 4294967296#
    For i = 0 To 3: buffer(64 + i) = unsigned - 256# * Fix(unsigned / 256#): unsigned = Fix(unsigned / 256#): Next
    For i = 1 To Len(documentID): buffer(67 + i) = AscW(Mid$(documentID, i, 1)): Next
    If revision >= 4 And Not metadata Then
        For i = count - 4 To count - 1: buffer(i) = 255: Next
    End If
    size = bits \ 8: If revision = 2 Then size = 5
    digestBytes = MD5(buffer, count)
    If revision >= 3 Then
        For i = 1 To 50: digestBytes = MD5(digestBytes, size): Next
    End If
    ReDim key(0 To size - 1): For i = 0 To size - 1: key(i) = digestBytes(i): Next
    If revision = 2 Then
        ReDim buffer(0 To 31): For i = 0 To 31: buffer(i) = pad(i): Next
        probe = RC4(buffer, key): count = 32
    Else
        ReDim buffer(0 To 31 + Len(documentID))
        For i = 0 To 31: buffer(i) = pad(i): Next
        For i = 1 To Len(documentID): buffer(31 + i) = AscW(Mid$(documentID, i, 1)): Next
        digestBytes = MD5(buffer, UBound(buffer) + 1): probe = RC4(digestBytes, key)
        ReDim roundKey(0 To size - 1)
        For i = 1 To 19
            For j = 0 To size - 1: roundKey(j) = key(j) Xor i: Next
            probe = RC4(probe, roundKey)
        Next
        count = 16
    End If
    For i = 0 To count - 1
        If probe(i) <> AscW(Mid$(user, i + 1, 1)) Then Err.Raise 5, "modPDFCrypto.PDFEmptyPasswordKey", "Password-protected PDF: local import supports forms that open without a password."
    Next
    PDFEmptyPasswordKey = key
End Function
Public Function PDFDecryptObject(ByRef source() As Byte, ByRef fileKey() As Byte, ByVal objectID As Long, ByVal generation As Long, ByVal method As String) As Byte()
    Dim buffer() As Byte, keyBytes() As Byte, digestBytes() As Byte, result() As Byte, iv(0 To 15) As Byte, blob(0 To 27) As Byte
    Dim count As Long, size As Long, i As Long, mode As Long, provider As LongPtr, key As LongPtr, failure As Long, message As String
    On Error GoTo Failed
    If method = "Identity" Then PDFDecryptObject = source: Exit Function
    size = UBound(fileKey) + 1: count = size + 5: If method = "AESV2" Then count = count + 4
    ReDim buffer(0 To count - 1)
    For i = 0 To size - 1: buffer(i) = fileKey(i): Next
    For i = 0 To 2: buffer(size + i) = objectID Mod 256: objectID = objectID \ 256: Next
    buffer(size + 3) = generation Mod 256: buffer(size + 4) = generation \ 256
    If method = "AESV2" Then buffer(size + 5) = 115: buffer(size + 6) = 65: buffer(size + 7) = 108: buffer(size + 8) = 84
    digestBytes = MD5(buffer, count): size = size + 5: If size > 16 Then size = 16
    ReDim keyBytes(0 To size - 1): For i = 0 To size - 1: keyBytes(i) = digestBytes(i): Next
    If method = "V2" Then PDFDecryptObject = RC4(source, keyBytes): Exit Function
    If method <> "AESV2" Or size <> 16 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Unsupported PDF encryption filter."
    count = UBound(source) + 1
    If count < 32 Or count Mod 16 <> 0 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Invalid encrypted PDF packet length."
    For i = 0 To 15: iv(i) = source(i): Next
    count = count - 16: ReDim result(0 To count - 1)
    For i = 0 To count - 1: result(i) = source(i + 16): Next
    blob(0) = 8: blob(1) = 2: blob(4) = 14: blob(5) = 102: blob(8) = 16
    For i = 0 To 15: blob(12 + i) = keyBytes(i): Next
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Windows PDF decryption unavailable."
    If CryptImportKey(provider, blob(0), 28, 0, 0, key) = 0 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Cannot initialize PDF packet key."
    mode = 1
    If CryptSetKeyParam(key, 4, mode, 0) = 0 Or CryptSetKeyParam(key, 1, iv(0), 0) = 0 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Cannot initialize PDF packet cipher."
    If CryptDecrypt(key, 0, 1, 0, result(0), count) = 0 Or count = 0 Then Err.Raise 5, "modPDFCrypto.PDFDecryptObject", "Invalid encrypted PDF packet."
    ReDim Preserve result(0 To count - 1): PDFDecryptObject = result
    CryptDestroyKey key: CryptReleaseContext provider, 0: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    If key <> 0 Then CryptDestroyKey key
    If provider <> 0 Then CryptReleaseContext provider, 0
    Err.Raise failure, "PDFDecryptObject", message
End Function
