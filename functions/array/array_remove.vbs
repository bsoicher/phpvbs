'=======================================================================
'�z��̎w�肵���v�f����폜����
'=======================================================================
'�y�����z
'  mAry     = array  �ΏۂƂȂ�z��
'  num      = int    �폜����v�f�ԍ�
'�y�߂�l�z
'  ������̔z���Ԃ��܂��B
'�y�����z
'  �EmAry ��num�Ԗڂ̗v�f���폜���܂��B
'  �E�z�� mAry �̒����͈�������܂��B
'=======================================================================
Function array_remove(mAry,num)

    if Not isArray(mAry) Then Exit Function
    If Not isNumeric(num) Then Exit Function

    Dim strCount
    strCount = uBound(mAry)
    If strCount+1 < num Then
        array_remove = mAry
        Exit Function
    End If

    If (strCount+1) = num Then
        ReDim Preserve mAry(strCount - 1)
        array_remove = mAry
        Exit Function
    End If

    If num = 0 Then
        call array_shift(mAry)
        array_remove = mAry
        Exit Function
    End If

    Dim tmpAry,retAry
    tmpAry = array_chunk(mAry,num)
    call array_shift( tmpAry(1) )

    call array_push(tmpAry(0),tmpAry(1))
    retAry = tmpAry(0)

    if uBound(tmpAry) > 1 Then

        Dim intCounter
        For intCounter = 2 to uBound(tmpAry)
            call array_push(retAry,tmpAry(intCounter))
        Next

    end if

    array_remove = retAry

End Function
