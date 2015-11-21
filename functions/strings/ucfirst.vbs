'=======================================================================
' 文字列の最初の文字を大文字にする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・str  の最初の文字がアルファベットであれば、 それを大文字にします。
'=======================================================================
Function ucfirst(byVal str)

    Dim tmp
    tmp = left(str,1)
    tmp = Ucase(tmp)
    ucfirst = tmp & Mid(str,2)

End Function