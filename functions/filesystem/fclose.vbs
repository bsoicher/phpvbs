'=======================================================================
' オープンされたファイルポインタをクローズする
'=======================================================================
'【引数】
'  
'【戻り値】
'  
'【処理】
'  ファイルをクローズします。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fclose
    ts.close
    Set ts = Nothing
end function