'=======================================================================
' 文字列を大文字にする
'=======================================================================
'【引数】
'  str     = string    入力文字列。
'【戻り値】
'  大文字にした文字列を返します。
'【処理】
'  ・str  のアルファベット部分をすべて大文字にして返します。
'=======================================================================
Function strtoupper(str)
    strtoupper = Ucase(str)
End Function
