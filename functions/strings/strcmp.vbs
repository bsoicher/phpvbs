'=======================================================================
' バイナリセーフな文字列比較
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  str1  が str2  よりも小さければ < 0 を、str1 が str2 よりも大きければ > 0 を、 等しければ 0 を返します。
'【処理】
'  ・この比較は大文字小文字を区別することに注意してください。
'=======================================================================
Function strcmp(str1, str2)
    strcmp = StrComp(str1,str2,vbBinaryCompare)
End Function