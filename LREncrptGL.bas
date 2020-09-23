Attribute VB_Name = "LREncrptGL"
Function Encodegl(vtext As String)
Dim a, b, c, d
Dim i
Dim x
Dim length
Dim temp
length = Len(vtext)

For i = 1 To length
    x = x + 1
    a = Asc(Mid(vtext, i, 1))
    b = Oct(a)
    c = Chr(b)
Next i



End Function
