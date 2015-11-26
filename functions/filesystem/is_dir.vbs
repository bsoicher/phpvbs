'=======================================================================
' ファイルがディレクトリかどうかを調べる
'=======================================================================
'【引数】
'  filename = string    ファイルへのパス。filename  が相対パスの場合は、現在の作業ディレクトリからの相対パスとして処理します。
'【戻り値】
'  ファイルがが存在して、かつそれがディレクトリであれば TRUE、それ以外の場合は FALSE を返します。
'【処理】
'  指定したファイルがディレクトリかどうかを調べます。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function is_dir(filename)

    Dim fn
    is_dir = false
    fn = fileMapPath(filename)

    If fso.FolderExists(fn) Then is_dir = true

End Function