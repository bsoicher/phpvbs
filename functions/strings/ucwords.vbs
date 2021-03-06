'=======================================================================
' 文字列の各単語の最初の文字を大文字にする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・ 文字がアルファベットの場合、str  の各単語の最初の文字を大文字にしたものを返します。
'  ・ 単語の定義は、空白文字 (スペース、フォームフィード、改行、キャリッジリターン、 水平タブ、垂直タブ) の直後にあるあらゆる文字からなる文字列です。 
'=======================================================================
Function ucwords(str)
    ucwords = preg_replace_callback("/^(.)|￥s(.)/","Ucase",str,"","")
End Function