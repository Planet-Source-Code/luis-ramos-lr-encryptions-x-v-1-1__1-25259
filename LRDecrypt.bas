Attribute VB_Name = "LRDecrypt"
Public Function DeCode(vText As String)
    Dim i As Integer
    Dim lrLen As Integer
    Dim lrChr As String
    Dim lrFin As String
    i = i + 1
    lrLen = Len(vText)


    Do While i <= lrLen


        DoEvents
            lrChr = Mid(vText, i, 3)


            Select Case lrChr
                
                Case "coe"
                lrChr = "a"
                Case "wer"
                lrChr = "b"
                Case "ibq"
                lrChr = "c"
                Case "am7"
                lrChr = "d"
                Case "pm1"
                lrChr = "e"
                Case "mop"
                lrChr = "f"
                Case "9v4"
                lrChr = "g"
                Case "qu6"
                lrChr = "h"
                Case "zxc"
                lrChr = "i"
                Case "4mp"
                lrChr = "j"
                Case "f88"
                lrChr = "k"
                Case "qe2"
                lrChr = "l"
                Case "vbn"
                lrChr = "m"
                Case "qwt"
                lrChr = "n"
                Case "pl5"
                lrChr = "o"
                Case "13s"
                lrChr = "p"
                Case "c%l"
                lrChr = "q"
                Case "w$w"
                lrChr = "r"
                Case "6a@"
                lrChr = "s"
                Case "!2&"
                lrChr = "t"
                Case "(=c"
                lrChr = "u"
                Case "wvf"
                lrChr = "v"
                Case "dp0"
                lrChr = "w"
                Case "w$-"
                lrChr = "x"
                Case "vn&"
                lrChr = "y"
                Case "c*4"
                lrChr = "z"
                
                
                Case "aq@"
                lrChr = "1"
                Case "902"
                lrChr = "2"
                Case "2.&"
                lrChr = "3"
                Case "/w!"
                lrChr = "4"
                Case "|pq"
                lrChr = "5"
                Case "ml|"
                lrChr = "6"
                Case "t'?"
                lrChr = "7"
                Case ">^s"
                lrChr = "8"
                Case "<s^"
                lrChr = "9"
                Case ";&c"
                lrChr = "0"
                
                
                Case "$)c"
                lrChr = "A"
                Case "-gt"
                lrChr = "B"
                Case "|p*"
                lrChr = "C"
                Case "1" & Chr(34) & "r"
                lrChr = "D"
                Case "c>:"
                lrChr = "E"
                Case "@+x"
                lrChr = "F"
                Case "v^a"
                lrChr = "G"
                Case "]eE"
                lrChr = "H"
                Case "aP0"
                lrChr = "I"
                Case "{=1"
                lrChr = "J"
                Case "cWv"
                lrChr = "K"
                Case "cDc"
                lrChr = "L"
                Case "*,!"
                lrChr = "M"
                Case "fW" & Chr(34)
                lrChr = "N"
                Case ".?T"
                lrChr = "O"
                Case "%<8"
                lrChr = "P"
                Case "@:a"
                lrChr = "Q"
                Case "&c$"
                lrChr = "R"
                Case "WnY"
                lrChr = "S"
                Case "{Sh"
                lrChr = "T"
                Case "_%M"
                lrChr = "U"
                Case "}'$"
                lrChr = "V"
                Case "QlU"
                lrChr = "W"
                Case "Im^"
                lrChr = "X"
                Case "l|P"
                lrChr = "Y"
                Case ".>#"
                lrChr = "Z"
                
                Case "\" & Chr(34) & "]"
                lrChr = "!"
                Case "cY,"
                lrChr = "@"
                Case "x%B"
                lrChr = "#"
                Case "a*v"
                lrChr = "$"
                Case "'&T"
                lrChr = "%"
                Case ";%R"
                lrChr = "^"
                Case "eG_"
                lrChr = "&"
                Case "Z/e"
                lrChr = "*"
                Case "rG\"
                lrChr = "("
                Case "]*F"
                lrChr = ")"
                Case "@B*"
                lrChr = "_"
                Case "+Hc"
                lrChr = "-"
                Case "&|D"
                lrChr = "="
                Case "(:#"
                lrChr = "+"
                Case "SlW"
                lrChr = "["
                Case "'QB"
                lrChr = "]"
                Case "{D>"
                lrChr = "{"
                Case "+c%"
                lrChr = "}"
                Case "(s:"
                lrChr = ":"
                Case "^a("
                lrChr = ";"
                Case "16."
                lrChr = "'"
                Case "s.*"
                lrChr = Chr(34)
                Case "&?W"
                lrChr = ","
                Case "GPQ"
                lrChr = "."
                Case "SK*"
                lrChr = "<"
                Case "RL^"
                lrChr = ">"
                Case "40C"
                lrChr = "/"
                Case "?#9"
                lrChr = "?"
                Case "_?/"
                lrChr = "\"
                Case "(_@"
                lrChr = "|"
                Case "=#B"
                lrChr = " "
            End Select
        lrFin = lrFin & lrChr
        i = i + 3


        DoEvents
        Loop
        DeCode = lrFin
    End Function
