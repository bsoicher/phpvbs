'=======================================================================
' fwrite() のエイリアス
'=======================================================================
'【引数】
'  str  =  string 書き込む文字列。
'【説明】
'  この関数は次の関数のエイリアスです。 fwrite().
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fputs(str)
    fputs = fwrite(str)
end function