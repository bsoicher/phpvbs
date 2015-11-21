'=======================================================================
' パス中のファイル名の部分を返す
'=======================================================================
'【引数】
'  path      = string   パス。
'  suffix    = string   ファイル名が、 suffix  で終了する場合、 この部分もカットされます。
'【戻り値】
'  指定した path  のベース名を返します。
'【処理】
'  ・この関数は、ファイルへのパスを有する文字列を引数とし、 ファイルのベース名を返します。
'=======================================================================
Function basename(path, suffix)

    Dim b
    b = preg_replace("/^.*[￥/￥￥]/g","",path,null,null)

    If len(suffix) > 0 Then
        If Right(b,len(suffix)) = suffix Then
            b = Left(b,len(b) - len(suffix))
        End If
    End If

    basename = b

End Function