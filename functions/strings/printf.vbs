'=======================================================================
' フォーマット済みの文字列を出力する
'=======================================================================
'【引数】
'  format   = string フォーマット文字列
'  args     = mixed  数値や文字列
'【戻り値】
'  フォーマット文字列 format  に基づき生成された文字列を出力します。
'【処理】
'  ・format  にしたがって、出力を生成します。
'=======================================================================
Function printf( format , args)
    Response.Write sprintf(format,args)
End Function