'=======================================================================
' 通常ファイルかどうかを調べる
'=======================================================================
'【引数】
'  filename = string    ファイルへのパス。
'【戻り値】
'  ファイルが存在し、かつそれが通常のファイルである場合に TRUE、 それ以外の場合に FALSE を返します。
'【処理】
'  指定したファイルが通常のファイルかどうかを調べます。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public Function is_file(ByVal filename)

    is_file = false
    filename = fileMapPath(filename)
    If fso.FileExists(filename) Then is_file = true

End Function