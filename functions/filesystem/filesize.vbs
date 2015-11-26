'=======================================================================
' ファイルのサイズを取得する
'=======================================================================
'【引数】
'  filename = string   ファイルへのパス。
'【戻り値】
'  ファイルのサイズを返し、エラーの場合は FALSE を返します (また E_WARNING レベルのエラーを発生させます) 。
'【処理】
'  指定したファイルのサイズを取得します。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function filesize(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filesize = f.Size

End Function