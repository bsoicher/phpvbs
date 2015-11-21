'=======================================================================
' ファイルパスに関する情報を返す
'=======================================================================
'【引数】
'  path     = string    調べたいパス。
'  options  = string    どの要素を返すのかをオプションのパラメータ options  で指定します。これは PATHINFO_DIRNAME、 PATHINFO_BASENAME、 PATHINFO_EXTENSION および PATHINFO_FILENAME の組み合わせとなります。 デフォルトではすべての要素を返します。
'【戻り値】
'   以下の要素を含む連想配列を返します。 dirname (ディレクトリ名)、basename (ファイル名) そして、もし存在すれば extension (拡張子)。
'   options を使用すると、 すべての要素を選択しない限りこの関数の返り値は文字列となります。 
'【処理】
'  pathinfo() は、path  に関する情報を有する連想配列を返します。
' File_System class(http://phpvbs.verygoodtown.com/)内での処理
'=======================================================================

Const PATHINFO_DIRNAME = 1
Const PATHINFO_BASENAME = 2
Const PATHINFO_EXTENSION = 4
Const PATHINFO_FILENAME = 3

Public Function pathinfo(ByVal path,ByVal options)

    Dim obj : set obj = Server.CreateObject("Scripting.Dictionary")
    Dim tmp


    obj("dirname") = dirname(path)
    obj("basename") = basename(path,"")
    obj("extension") = obj("basename")

    If inStr(obj("basename"),".") Then
        tmp = Split(obj("basename"),".")
        obj("extension") = tmp( uBound(tmp) )
    End if

    obj("filename") = Replace(obj("basename"),"." & obj("extension"),"")

    If len(options) > 0 Then

        If options = PATHINFO_DIRNAME Then
            pathinfo = obj("dirname")
        ElseIf options = PATHINFO_BASENAME Then
            pathinfo = obj("basename")
        ElseIf options = PATHINFO_EXTENSION Then
            pathinfo = obj("extension")
        ElseIf options = PATHINFO_FILENAME Then
            pathinfo = obj("filename")
        End if
        Exit Function
    End If

    set pathinfo = obj
End Function