'=======================================================================
' fwrite() のエイリアス
'=======================================================================
'【引数】
'  str  =  string 書き込む文字列。
'【説明】
'  この関数は次の関数のエイリアスです。 fwrite().
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public function fputs(str)
    fputs = fwrite(str)
end function