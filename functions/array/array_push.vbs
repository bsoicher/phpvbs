'=======================================================================
'��ȏ�̗v�f��z��̍Ō�ɒǉ�����
'=======================================================================
'�y�����z
'  mAry     = array  �z��
'  mVal     = mixed  �ǉ�����v�f
'�y�߂�l�z
'  ������̔z��̒��̗v�f�̐���Ԃ��܂��B
'�y�����z
'  �E�n���ꂽ�ϐ��� mAry �̍Ō�ɉ����܂��B
'  �E�z�� mAry �̒����͓n���ꂽ�ϐ��̐������������܂��B
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