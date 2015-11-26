'=======================================================================
' ディレクトリを作る
'=======================================================================
'【引数】
'  pathname = string    ディレクトリのパス。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  指定したディレクトリを作成します。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function mkdir(ByVal pathname)

    mkdir = false
    pathname = fileMapPath(pathname)
    If not fso.FolderExists(pathname) Then
        mkdir = fso.CreateFolder(pathname)
    End If

End Function