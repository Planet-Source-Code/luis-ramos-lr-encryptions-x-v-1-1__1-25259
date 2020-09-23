Attribute VB_Name = "LREncryptHL"
Function Encodehl(vText As String) As String
    Dim a As String, b As String, x As Long, s() As Byte
    Dim c As String, d As String
    If Not Len(vText) Mod 2 = 0 Then vText = vText & " "
    ReDim s(Len(vText) - 1)
    For x = 1 To Len(vText)
        a = Left(Right("0" & Hex(Asc(Mid(vText, x, 1))), 2), 1)
        b = Right(Right("0" & Hex(Asc(Mid(vText, x, 1))), 2), 1)
        c = Left(Right("0" & Hex(Asc(Mid(vText, Len(vText) - x + 1, 1))), 2), 1)
        d = Right(Right("0" & Hex(Asc(Mid(vText, Len(vText) - x + 1, 1))), 2), 1)
        s(x - 1) = Val("&H" & a & c)
        s(Len(vText) - x) = Val("&H" & b & d)
    Next
    Encodehl = StrConv(s, vbUnicode)
End Function

