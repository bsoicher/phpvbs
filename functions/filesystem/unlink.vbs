'=======================================================================
' ファイルを削除する
'=======================================================================
'【引数】
'  filename     = string    ファイルへのパス。
'【戻り値】
'   成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  filename  を削除します。 Unix C 言語の関数 unlink() と動作は同じです。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function unlink(ByVal filename)

    filename = fileMapPath(filename)
    fso.DeleteFile filename
    unlink = true

End Function