'=======================================================================
' 平方根
'=======================================================================
'【引数】
'  arg = float  処理する引数。
'【戻り値】
'  arg  の平方根を返します。 負の数を指定した場合は、null を返します。
'【処理】
'  ・arg  の平方根を返します。
'=======================================================================
Function sqrt(arg)

    If arg < 0 Then
        sqrt = null
    Else
        sqrt = sqr(arg)
    End If

End Function