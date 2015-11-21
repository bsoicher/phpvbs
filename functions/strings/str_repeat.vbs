'=======================================================================
' •¶š—ñ‚ğ”½•œ‚·‚é
'=======================================================================
'yˆø”z
'  input        = string    ŒJ‚è•Ô‚·•¶š—ñB
'  multiplier   = int       input ‚ğŒJ‚è•Ô‚·‰ñ”Bmultiplier ‚Í 0 ˆÈã‚Å‚È‚¯‚ê‚Î‚È‚è‚Ü‚¹‚ñB multiplier ‚ª 0 ‚Éİ’è‚³‚ê‚½ê‡A‚±‚ÌŠÖ”‚Í‹ó•¶š‚ğ•Ô‚µ‚Ü‚·B
'y–ß‚è’lz
'  ŒJ‚è•Ô‚µ‚½•¶š—ñ‚ğ•Ô‚µ‚Ü‚·B
'yˆ—z
'  Einput  ‚ğ multiplier  ‰ñ‚ğŒJ‚è•Ô‚µ‚½•¶š—ñ‚ğ•Ô‚µ‚Ü‚·B
'=======================================================================
Function str_repeat(input, multiplier)
    If multiplier < 0 Then Exit Function
    str_repeat = String(multiplier,input)
End Function