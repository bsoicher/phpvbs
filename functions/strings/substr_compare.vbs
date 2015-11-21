'=======================================================================
' w’è‚µ‚½ˆÊ’u‚©‚çw’è‚µ‚½’·‚³‚Ì 2 ‚Â‚Ì•¶š—ñ‚É‚Â‚¢‚ÄAƒoƒCƒiƒŠ‘Î‰‚Å”äŠr‚·‚é
'=======================================================================
'yˆø”z
'  main_str             = string    Å‰‚Ì•¶š—ñB
'  str                  = string    Ÿ‚Ì•¶š—ñB
'  offset               = int       ”äŠr‚ğŠJn‚·‚éˆÊ’uB •‰‚Ì’l‚ğw’è‚µ‚½ê‡‚ÍA•¶š—ñ‚ÌÅŒã‚©‚ç”‚¦‚Ü‚·B
'  length               = int       ”äŠr‚·‚é’·‚³B
'  case_insensitivity   = bool      case_insensitivity  ‚ª TRUE ‚Ìê‡A ‘å•¶š¬•¶š‚ğ‹æ•Ê‚¹‚¸‚É”äŠr‚µ‚Ü‚·B
'y–ß‚è’lz
'  main_str  ‚Ì offset  ˆÈ~‚ª str  ‚æ‚è¬‚³‚¢ê‡‚É•‰‚Ì”A str  ‚æ‚è‘å‚«‚¢ê‡‚É³‚Ì”A “™‚µ‚¢ê‡‚É 0 ‚ğ•Ô‚µ‚Ü‚·B
'yˆ—z
'  E substr_compare() ‚ÍAmain_str  ‚Ì offset  •¶š–ÚˆÈ~‚ÌÅ‘å length  •¶š‚ğAstr  ‚Æ”äŠr‚µ‚Ü‚·B
'=======================================================================
Function substr_compare(main_str,str,offset,length, case_insensitivity)

    If len(offset) > 0 Then
        If offset > 0 Then
            main_str = Mid(main_str,offset)
        Else
            main_str = Mid(main_str,len(main_str) + offset + 1)
        End If
    End If

    If len(length) > 0 Then
        main_str = Left(main_str,length)
        str = Left(str,length)
    End If
    var_dump main_str
    var_dump str
    If case_insensitivity = true Then
        substr_compare = strcasecmp(main_str,str)
    Else
        substr_compare = strcmp(main_str,str)
    End If

End Function