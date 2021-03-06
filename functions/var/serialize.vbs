'=======================================================================
'値の保存可能な表現を生成する
'=======================================================================
'【引数】
'  val   = mixed    シリアル化する値。
'【戻り値】
'  val  の保存可能なバイトストリーム表現を含む文字列を返します。 
'【処理】
'  ・ 値の保存可能な表現を生成します。
'  ・ 型や構造を失わずに ASP の値を保存または渡す際に有用です。
'  ・ シリアル化された文字列を ASP の値に戻すには、 unserialize() を使用してください。 
'=======================================================================
Function serialize(ByVal val)

    Dim strstrType
    strType = getType(val)

    Dim str
    Dim cnt : cnt = 0
    Dim strs
    Dim key

    Select Case strType

        Case "vbEnpty","vbNull"
            str = "N"
        Case "vbBoolean"
            str = "b:" & [?](val,1,0)
        Case "vbInteger","vbLong","vbSingle","vbDouble","vbCurrency"
            str = [?]([==](int(val),val),"i","d") & ":" & val
        Case "vbDate","vbString","vbVariant"
            str = "s:" & len(val) & ":""" & val & """"
        Case "vbArray"
            str = "a"

            For key = 0 to uBound(val)
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"
            str = str & ";"

        Case "vbObject"
            str = "O"

            For Each key In val
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"

        Case Else
            'empty
    End Select

    If strType <> "vbArray" AND strType <> "vbObject" Then
        str = str & ";"
    End If

    serialize = str
End Function