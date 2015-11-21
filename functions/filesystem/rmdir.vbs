'=======================================================================
' ディレクトリを削除する
'=======================================================================
'【引数】
'  dirname     = string    ディレクトリへのパス。
'【戻り値】
'   成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  dirname で指定されたディレクトリを 削除しようと試みます。
'  ディレクトリは空でなくてはならず、また 適切なパーミッションが設定されていなければなりません。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public Function rmdir(ByVal dirname)

    dirname = fileMapPath(dirname)
    fso.DeleteFolder dirname
    rmdir = true

End Function