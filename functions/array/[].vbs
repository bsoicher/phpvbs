'=======================================================================
'��̗v�f��z��̍Ō�ɒǉ�����
'=======================================================================
'�y�����z
'  mAry     = mixed  �z��
'  mVal     = mixed  �ǉ�����v�f
'�y�߂�l�z
'  �l��Ԃ��܂���B
'�y�����z
'  �E�n���ꂽ�ϐ��� mAry  �̍Ō�ɉ����܂��B
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