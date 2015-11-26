'=======================================================================
' ファイルまたは URL をオープンする
'=======================================================================
'【引数】
'  filename  =  string データを読み込みたいファイルの名前。
'  mode      =  string ストリームに要するアクセス形式を指定します
'【戻り値】
'  成功した場合にファイルポインタリソースを返します。
'【処理】
'  fopen() は、filename  で指定されたリソースをストリームに結び付けます。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fopen(filename, mode)

    Dim filePath
    filePath = fileMapPath(filename)

    If left(filePath,len("http://")) = "http://" Then
        fopen = file_get_contents(filePath)
        Exit Function
    End If

    Select Case mode
    Case "r"
        '読み込みのみでオープンします。
        Set ts = fso.OpenTextFile(filePath,1,false)
    Case "w"
        '書き込みでオープンします。
        Set ts = fso.OpenTextFile(filePath,2,true)
    Case "a"
        '追記でオープンします。
        Set ts = fso.OpenTextFile(filePath,8,true)
    Case "x"
        '書き込みでオープンします。ファイルが存在した場合はfalseを返します。
        If is_file(filePath) Then
            fopen = false
        Else
            Set ts = fso.OpenTextFile(filePath,2,true)
        End If
    Case Else
        'empty
        ts = false
    End Select

end function