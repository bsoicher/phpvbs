'=======================================================================
' ファイルの更新時刻を取得する
'=======================================================================
'【引数】
'  filename = string   ファイルへのパス。
'【戻り値】
'  ファイルの最終更新時刻を返し、エラーの場合は FALSE  を返します。
'【処理】
'  この関数は、ファイルのブロックデータが書き込まれた時間を返します。 これは、ファイルの内容が変更された際の時間です。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public Function filemtime(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filemtime = f.DateLastModified

End Function

'*************************************
Private Function fileMapPath(filename)

    Dim tmp
    tmp = Left(filename,3)
    tmp = Lcase(tmp)
    If tmp <> "d:￥" and tmp <> "c:￥" and left(filename,7) <> "http://" then
            fileMapPath = Server.MapPath(filename)
    Else
        fileMapPath = filename
    End If

End Function