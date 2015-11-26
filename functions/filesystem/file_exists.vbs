'=======================================================================
' ファイルまたはディレクトリが存在するかどうか調べる
'=======================================================================
'【引数】
'  path      = string   ファイルあるいはディレクトリへのパス。
'【戻り値】
'  ファイルあるいはディレクトリが存在するかどうかを調べます。
'【処理】
'  ファイルあるいはディレクトリが存在するかどうかを調べます。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function file_exists(ByVal filename)

    file_exists = false
    filename = fileMapPath(filename)
    If fso.FileExists(filename) Then file_exists = true
    If fso.FolderExists(filename) Then file_exists = true

End Function
