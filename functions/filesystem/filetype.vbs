'=======================================================================
' ファイルタイプを取得する
'=======================================================================
'【引数】
'  filename = string   ファイルへのパス。
'【戻り値】
'  ファイルのタイプを返します。
'【処理】
'  指定したファイルのタイプを返します。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function filetype(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filetype = f.Type

End Function