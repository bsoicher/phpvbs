'=======================================================================
' •¶š—ñ‚Ì’†‚©‚ç”CˆÓ‚Ì•¶š‚ğ’T‚·
'=======================================================================
'yˆø”z
'  haystack     =   string  char_list  ‚ğ’T‚·•¶š—ñB
'  char_list    =   string  ‚±‚Ìƒpƒ‰ƒ[ƒ^‚Í‘å•¶š¬•¶š‚ğ‹æ•Ê‚µ‚Ü‚·B
'y–ß‚è’lz
'  Œ©‚Â‚©‚Á‚½•¶š‚©‚çn‚Ü‚é•¶š—ñA‚ ‚é‚¢‚ÍŒ©‚Â‚©‚ç‚È‚©‚Á‚½ê‡‚É FALSE ‚ğ•Ô‚µ‚Ü‚·B
'yˆ—z
'  E strpbrk() ‚ÍA•¶š—ñ haystack  ‚©‚ç char_list  ‚ğ’T‚µ‚Ü‚·B
'=======================================================================
Function strpbrk( haystack, char_list )

    haystack  = Cstr( haystack )
    char_list = Cstr( char_list )

    Dim intlen : intlen = len( haystack )
    Dim i,char
    For i = 1 to intlen
        char = Mid(haystack,i,1)
        If [!==](strpos(char_list,char,""),false) Then
            strpbrk = Mid(haystack,i)
            Exit Function
        End If
    Next

    strpbrk = false

End Function