'=======================================================================
' バイナリセーフなファイル書き込み処理
'=======================================================================
'【引数】
'  str     =  string   書き込む文字列。
'【戻り値】
'  
'【処理】
'  string の内容を ファイル・ストリームに書き込みます。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================
Public function fwrite(str)
    ts.WriteLine str
end function