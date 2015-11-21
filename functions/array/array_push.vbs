'=======================================================================
'ˆê‚ÂˆÈã‚Ì—v‘f‚ğ”z—ñ‚ÌÅŒã‚É’Ç‰Á‚·‚é
'=======================================================================
'yˆø”z
'  mAry     = array  ”z—ñ
'  mVal     = mixed  ’Ç‰Á‚·‚é—v‘f
'y–ß‚è’lz
'  ˆ—Œã‚Ì”z—ñ‚Ì’†‚Ì—v‘f‚Ì”‚ğ•Ô‚µ‚Ü‚·B
'yˆ—z
'  E“n‚³‚ê‚½•Ï”‚ğ mAry ‚ÌÅŒã‚É‰Á‚¦‚Ü‚·B
'  E”z—ñ mAry ‚Ì’·‚³‚Í“n‚³‚ê‚½•Ï”‚Ì”‚¾‚¯‘‰Á‚µ‚Ü‚·B
'=======================================================================
Function array_push(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            intElementCount = UBound(mAry)
            ReDim Preserve mAry(intElementCount + UBound(mVal) + 1)

            For intCounter = 0 to UBound(mVal)
                mAry(intElementCount + intCounter + 1) = mVal(intCounter)
            Next

        Else
            ReDim Preserve mAry(UBound(mAry) + 1)
            mAry(UBound(mAry)) = mVal
        End If
    Else

        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_push = UBound(mAry)

End Function