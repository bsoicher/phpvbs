'=======================================================================
'�w�蒷�A�w�肵���l�Ŕz��𖄂߂�
'=======================================================================
'�y�����z
'  mAry         = array  �l�𖄂߂���ƂƂȂ�z��B
'  pad_size     = int    �V�����z��̃T�C�Y�B
'  pad_value    = mixed  mAry �� pad_size ��菬�����Ƃ��ɁA ���߂邽�߂Ɏg�p����l�B
'�y�߂�l�z
'  pad_size  �Ŏw�肵�������ɂȂ�悤�ɒl pad_value  �Ŗ��߂� mAry �̃R�s�[��Ԃ��܂��B
'  pad_size  �����̏ꍇ�A�z��̉E�������߂��܂��B 
'  ���̏ꍇ�A�z��̍��������߂��܂��B 
'  pad_size  �̐�Βl�� mAry �̒����ȉ��̏ꍇ�A���߂鏈���͍s���܂���B
'�y�����z
'  pad_size  �Ŏw�肵�������ɂȂ�悤�ɒl pad_value  �Ŗ��߂� mAry �̃R�s�[��Ԃ��܂��B
'=======================================================================
Function array_pad(ByVal mAry, pad_size, pad_value)

    If Not isArray( mAry ) Then Exit Function
    If Not isNumeric( pad_size ) Then Exit Function

    Dim pad,aryCounter,newLength,i,intCounter

    If pad_size < 0 Then
        newLength = pad_size * -1
    Else
        newLength = pad_size
    End If
    newLength = newLength -1

    aryCounter = uBound(mAry)
    If newLength > aryCounter Then

        ReDim pad(newLength)
        intCounter = 0
        For i = 0 to newLength
            If pad_size < 0 Then
                If newLength - aryCounter > i Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(intCounter)
                    intCounter = intCounter + 1
                End If
            Else
                If i > aryCounter Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(i)
                End If
            End If
        Next
    Else
        pad = mAry
    End If

    array_pad = pad

End Function