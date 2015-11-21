'=======================================================================
' 指定したファイルのMD5ハッシュ値を計算する
'=======================================================================
'【引数】
'  filename = string    ファイル名
'【戻り値】
'  成功時は文字列、そうでなければ FALSE
'【処理】
'  ・filename パラメータで指定したファイルの MD5ハッシュを計算し、そのハッシュを返します。
'=======================================================================
Function md5_file(filename)
    md5_file = md5( file_get_contents(filename) )
End Function