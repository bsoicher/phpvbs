'=======================================================================
' ファイルポインタから 1 行取り出し、HTML タグを取り除く
'=======================================================================
'【引数】
'  
'【戻り値】
'  HTML や PHP コードを取り除いた文字列を返します。
'【処理】
'  fgets() と同じですが、 fgetss() は読み込んだテキストから HTML および PHP のタグを取り除こうとすることが異なります。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fgetss
    fgets = strip_tags(ts.ReadLine)
end function