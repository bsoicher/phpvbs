'=======================================================================
' ファイルポインタから1文字取り出す
'=======================================================================
'【引数】
'  
'【戻り値】
'  ファイルポインタから 1 文字読み出し、 その文字からなる文字列を返します。EOF の場合に FALSE を返します。
'【処理】
'  指定したファイルポインタから 1 文字読み出します。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fgetc
    If ts.AtEndofStream Then
        fgetc = false
    Else
        fgetc = ts.Read(1)
    End If
end function