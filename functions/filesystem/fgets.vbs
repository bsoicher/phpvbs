'=======================================================================
' ファイルポインタから 1 行取得する
'=======================================================================
'【引数】
'  
'【戻り値】
'  ファイルポインタから 1 行取得する
'【処理】
'  ファイルポインタから 1 行取得します。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fgets
    fgets = ts.ReadLine
end function