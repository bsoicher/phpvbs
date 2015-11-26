'=======================================================================
'ファイルの内容を全て文字列に読み込む
'=======================================================================
'【引数】
'  filename  = string データを読み込みたいファイルの名前。
'【戻り値】
'  読み込んだデータを返します。失敗した場合は FALSE を返します。
'【処理】
'  ファイルの内容を文字列に読み込む
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function file_get_contents(filename)

    Dim ts
    Dim contents

    if left(filename,7) <> "http://" and file_exists( filename ) then
        Set TS = fso.OpenTextFile( fileMapPath(filename),1)

        '空のファイルの場合、エラーになってしまう
        If TS.AtEndOfStream <> True Then
           contents = TS.ReadAll
        End If

        file_get_contents = contents
        Exit Function
    end if

    if left(filename,7) <> "http://" then
        file_get_contents = false
        Exit Function
    end if

    Dim objWinHttp
    'Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
    'Set objWinHttp = Server.CreateObject("MSXML2.XMLHTTP")
    Set objWinHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
'        on error resume next

    objWinHttp.Open "GET", filename, false
    objWinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objWinHttp.Send
    'Response.Write objWinHttp.Status & " " & objWinHttp.StatusText

    file_get_contents = objWinHttp.ResponseText

    Set objWinHttp = Nothing

    '有効にしたエラー処理を無効にする
'        on error goto 0

end function