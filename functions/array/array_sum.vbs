'=======================================================================
'�z��̒��̒l�̍��v���v�Z����
'=======================================================================
'�y�����z
'  mAry         = Array ���͂̔z��B
'�y�߂�l�z
'  �E�l�̍��v�𐮐��܂��� float �Ƃ��ĕԂ��܂��B
'�y�����z
'  �E�z��̒��̒l�̍��v�𐮐��܂��� float �Ƃ��ĕԂ��܂��B
'=======================================================================
Function array_sum(mAry)

    array_sum = 0
    If Not isArray(mAry) and Not isObject(mAry) Then Exit Function

    Dim key
    If isObject(mAry) Then
        For Each key in mAry
            array_sum = array_sum + mAry(key)
        Next
    Else
        For Each key in mAry
            array_sum = array_sum + key
        Next
    End If

End Function