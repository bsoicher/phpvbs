'=======================================================================
' 文字列の長さを得る
'=======================================================================
'【引数】
'  str  = string    長さを調べる文字列。
'【戻り値】
'  成功した場合に str の長さ、 str  が空の文字列だった場合に 0 を返します。
'【処理】
'  ・与えられた str の長さを返します。
'=======================================================================
Function strlen(str)
    strlen = len(str)
End Function