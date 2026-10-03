Attribute VB_Name = "modFileHash"
Option Explicit
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal parameter As Long, ByRef data As Byte, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long
Public Function SourceFileHash(ByVal path As String) As String
    Dim provider As LongPtr, hash As LongPtr, buffer() As Byte, digest(0 To 31) As Byte
    Dim file As Integer, remaining As Long, length As Long, i As Long, opened As Boolean, failure As Long, message As String
    On Error GoTo Failed
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise vbObjectError + 600, , "Cannot initialize source hashing."
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then Err.Raise vbObjectError + 601, , "Cannot create source hash."
    file = FreeFile: Open path For Binary Access Read Shared As #file: opened = True
    remaining = LOF(file)
    Do While remaining > 0
        length = remaining: If length > 131072 Then length = 131072
        ReDim buffer(0 To length - 1)
        Get #file, , buffer
        If CryptHashData(hash, buffer(0), length, 0) = 0 Then Err.Raise vbObjectError + 602, , "Cannot hash source bytes."
        remaining = remaining - length
    Loop
    Close #file: opened = False
    length = 32
    If CryptGetHashParam(hash, 2, digest(0), length, 0) = 0 Then Err.Raise vbObjectError + 603, , "Cannot read source hash."
    For i = 0 To 31: SourceFileHash = SourceFileHash & Right$("0" & Hex$(digest(i)), 2): Next
    CryptDestroyHash hash: CryptReleaseContext provider, 0
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If opened Then Close #file
    If hash <> 0 Then CryptDestroyHash hash
    If provider <> 0 Then CryptReleaseContext provider, 0
    On Error GoTo 0
    Err.Raise failure, "SourceFileHash", message
End Function
