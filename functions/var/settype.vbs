'=======================================================================
'•Ï”‚ÌŒ^‚ğƒZƒbƒg‚·‚é
'=======================================================================
'yˆø”z
'  val   = mixed    ”jŠü‚·‚é•Ï”B
'  type  = string   type  ‚Ì’l‚ÍˆÈ‰º‚Ì–½—ß‚Ì‚¢‚¸‚ê‚©‚Å‚·B
'y–ß‚è’lz
'  ¬Œ÷‚µ‚½ê‡‚É TRUE ‚ğA¸”s‚µ‚½ê‡‚É FALSE ‚ğ•Ô‚µ‚Ü‚·B
'yˆ—z
'  E•Ï” str ‚ÌŒ^‚ğ type  ‚ÉƒZƒbƒg‚µ‚Ü‚·B
'=======================================================================
Function settype(ByRef str,strType)

    settype = true

    Select Case strType
    Case "bool"
        str = CBool(str)
    Case "boolean"
        str = CBool(str)
    Case "byte"
        str = CByte(str)
    Case "currency"
        str = CCur(str)
    Case "date"
        str = CDate(str)
    Case "float"
        str = CDbl(str)
    Case "double"
        str = CDbl(str)
    Case "int"
        str = Cint(str)
    Case "integer"
        str = Cint(str)
    Case "long"
        str = CLng(str)
    Case "single"
        str = CSng(str)
    Case "string"
        str = Cstr(str)
    Case "array"
        If not isArray(str) Then
            str = array(str)
        End If
    Case "null"
        str = null
    Case Else
        settype = false
    End Select

End Function