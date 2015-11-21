'=======================================================================
'ˆê‚Â‚Ì—v‘f‚ğ”z—ñ‚ÌÅŒã‚É’Ç‰Á‚·‚é
'=======================================================================
'yˆø”z
'  mAry     = mixed  ”z—ñ
'  mVal     = mixed  ’Ç‰Á‚·‚é—v‘f
'y–ß‚è’lz
'  ’l‚ğ•Ô‚µ‚Ü‚¹‚ñB
'yˆ—z
'  E“n‚³‚ê‚½•Ï”‚ğ mAry  ‚ÌÅŒã‚É‰Á‚¦‚Ü‚·B
'=======================================================================
Sub [](ByRef mAry, ByVal mVal)

    If IsArray(mAry) Then
        Dim counter : counter = UBound(mAry) + 1
        ReDim Preserve mAry(counter)
        mAry(counter) = mVal
    Else
        mAry = Array(mVal)
    End If

End Sub