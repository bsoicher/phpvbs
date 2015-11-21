'=======================================================================
'文字列のmd5ハッシュ値を計算する
'=======================================================================
'【引数】
'  str      = string  文字列
'【戻り値】
'  32 文字の 16 進数からなるハッシュを返します。
'【処理】
'  ・str の MD5 ハッシュ値を計算し、 そのハッシュを返します。
'=======================================================================
Function md5(str)

    Dim bobj
    Set bobj = Server.CreateObject("basp21")
    md5 = bobj.MD5(str)

End Function