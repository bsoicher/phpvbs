'=======================================================================
' ファイルの最終アクセス時刻を取得する
'=======================================================================
'【引数】
'  filename = string   ファイルへのパス。
'【戻り値】
'  ファイルの最終アクセス時刻を返し、エラーの場合は FALSE を返します。
'【処理】
'  指定したファイルの最終アクセス時刻を取得します。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public Function fileatime(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    fileatime = f.DateLastAccessed

End Function