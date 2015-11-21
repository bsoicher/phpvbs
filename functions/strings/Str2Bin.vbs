'=======================================================================
' 文字列をバイナリ変換
'=======================================================================
'【引数】
'  str  = string    バイナリ変換したい文字列。
'【戻り値】
'  バイナリ変換された文字列を返します。
'【処理】
'  ・日本語対応
'=======================================================================
Function Str2Bin(strData)

    Dim strChar,strHex
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        strHex = CStr(Hex(Asc(strChar)))
        Select Case Len(strHex)
            Case 1 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 1)))
            Case 2 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
            Case 4 '2Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 3, 2)))
        End Select
    Next
End Function