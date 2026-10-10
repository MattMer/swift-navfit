Attribute VB_Name = "modRUADProfiles"
Option Explicit
Public Function RUADKeys() As Variant
    RUADKeys = Split("CUI|RUIC|RESERVE UNIT|RESERVE ACTIVITY AND LOC|ST|UIC|NRC|NRA|RPC|RATE/RANK|NAME|RFAS Asgmt|E/O|SELRES TYPE|BIN|RFAS|REQ DESG/REQ RTING|PNOBC/PNEC|DESG|NOBC/NEC|AUIC|TRUIC|UMUIC|MOB ACTIVITY|MOB BILLET|PRD|IMS|MAS (A)|MAS (M)|MAS (T)|ABSC|RBSC|ADDRESS|PHONE|NROWS Email", "|")
End Function
Public Function RUADPages() As Variant
    RUADPages = Split("0|0|0|0|0|0|0|0|0|1|1|2|1|1|2|2|2|2|1|1|2|2|2|3|3|2|3|3|3|3|3|3|4|4|4", "|")
End Function
Public Function RUADDecode(ByVal xml As Variant) As Object
    Dim d As Object, doc As Object, n As Object, key As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each key In RUADKeys(): d.Add CStr(key), Null: Next
    If Len(Nz(xml, "")) > 0 Then
        Set doc = CreateObject("MSXML2.DOMDocument.6.0"): doc.async = False
        If Not doc.LoadXML(CStr(xml)) Then Err.Raise vbObjectError + 7060, , "RUAD profile data could not be read. No changes saved."
        For Each n In doc.SelectNodes("/RUAD/Field")
            key = n.getAttribute("key")
            If d.Exists(CStr(key)) Then
                If n.getAttribute("null") = "1" Then d(CStr(key)) = Null Else d(CStr(key)) = n.Text
            End If
        Next
    End If
    Set RUADDecode = d
End Function
Public Function RUADEncode(ByVal values As Object, ByVal memberName As String, Optional ByVal provenance As Variant) As String
    Dim doc As Object, parent As Object, n As Object, key As Variant, v As Variant, supported As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set parent = doc.createElement("RUAD"): doc.appendChild parent
    Set supported = RUADDecode(Null)
    For Each key In values.Keys
        If Not supported.Exists(CStr(key)) Then Err.Raise vbObjectError + 7060, , "Unsupported RUAD field: " & CStr(key)
    Next
    For Each key In RUADKeys()
        v = Null
        If values.Exists(CStr(key)) Then v = values(CStr(key))
        If key = "NAME" Then v = memberName
        If Not IsNull(v) Then
            If Len(CStr(v)) > 4000 Then Err.Raise vbObjectError + 7060, , CStr(key) & ": use no more than 4,000 characters. No value is truncated."
        End If
        Set n = doc.createElement("Field"): n.setAttribute "key", CStr(key)
        If IsNull(v) Then n.setAttribute "null", "1" Else n.Text = CStr(v)
        parent.appendChild n
    Next
    If Not IsMissing(provenance) Then
        If Len(Nz(provenance, "")) > 0 Then
            Dim sourceDoc As Object, sourceNode As Object
            Set sourceDoc = CreateObject("MSXML2.DOMDocument.6.0"): sourceDoc.async = False
            If Not sourceDoc.LoadXML(CStr(provenance)) Then Err.Raise vbObjectError + 7060, , "Existing RUAD provenance cannot be read."
            Set sourceNode = sourceDoc.SelectSingleNode("/RUAD/Source")
            If Not sourceNode Is Nothing Then parent.appendChild sourceNode.cloneNode(True)
        End If
    End If
    RUADEncode = doc.xml
End Function
Public Function UXRUADPage(ByVal page As Long) As Boolean
    Dim f As Access.Form, keys As Variant, pages As Variant, i As Long
    Set f = Forms!frmUXMemberEditor: keys = RUADKeys(): pages = RUADPages()
    f.Controls("ruadPage" & page).SetFocus
    For i = 0 To UBound(keys)
        If keys(i) <> "NAME" Then
            f.Controls("ruad" & i).Visible = (CLng(pages(i)) = page)
            f.Controls("ruadLabel" & i).Visible = (CLng(pages(i)) = page)
        End If
    Next
    For i = 0 To 4: f.Controls("ruadPage" & i).FontBold = (i = page): Next
    UXRUADPage = True
End Function
Public Sub UXRUADLoad(ByVal xml As Variant)
    Dim values As Object, keys As Variant, i As Long
    Set values = RUADDecode(xml): keys = RUADKeys()
    For i = 0 To UBound(keys)
        If keys(i) <> "NAME" Then Forms!frmUXMemberEditor.Controls("ruad" & i) = values(keys(i))
    Next
    UXRUADPage 0
End Sub
Public Function UXRUADValues() As Object
    Dim d As Object, keys As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary"): keys = RUADKeys()
    For i = 0 To UBound(keys)
        If keys(i) = "NAME" Then d.Add CStr(keys(i)), Forms!frmUXMemberEditor!txtName Else d.Add CStr(keys(i)), Forms!frmUXMemberEditor.Controls("ruad" & i).Value
    Next
    Set UXRUADValues = d
End Function
