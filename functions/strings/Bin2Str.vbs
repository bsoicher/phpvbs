'=======================================================================
' バイナリデータを文字列に変換
'=======================================================================
'【引数】
'  str  = string    文字列に変換したいバイナリデータ
'【戻り値】
'  文字列を返します。
'【処理】
'  ・日本語対応
'=======================================================================
Function Bin2Str(byteData)

    Dim i,u,s,intChar : i = 1

    Bin2Str =""
    Do While i <= LenB(byteData)
        u = Hex(AscB(MidB(byteData, i, 1)))
        If ((CInt("&H" & u) >= &H81) And (CInt("&H" & u) <= &H9F)) _
            Or ((CInt("&H" & u) >= &HE0) And (CInt("&H" & u) <= &HFC)) Then 'Code Page 932
            l = Hex(AscB(MidB(byteData, i + 1, 1)))
            intChar = CInt("&H" & u & l)
            s = Chr(intChar)
            i = i + 2
        Else
            intChar = CInt("&H" & u)
            s = Chr(intChar)
            i = i + 1
        End If
        Bin2Str = Bin2Str & s
    Loop
End Function
