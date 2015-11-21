'=======================================================================
'•Ï”‚Ì®”‚Æ‚µ‚Ä‚Ì’l‚ğæ“¾‚·‚é
'=======================================================================
'yˆø”z
'  var = mixed •¶š—ñ
'y–ß‚è’lz
'  ®”
'yˆ—z
'  Evar  ‚Ì integer ‚Æ‚µ‚Ä‚Ì’l‚ğ•Ô‚µ‚Ü‚·B
'=======================================================================
Function intval(str)

    intval = 1
    If IsObject(str) or IsArray(str) Then Exit Function
    If str = true Then Exit Function

    intval = 0
    If is_empty(str) or Not isNumeric(str) Then Exit Function

    str = int(str)
    If str > 32767 Then
        intval = 32767
    Else
        intval = Cint(str)
    End If

End Function