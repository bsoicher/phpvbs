'=======================================================================
'�z��̖�������v�f����菜��
'=======================================================================
'�y�����z
'  mAry = array  �l�����o���z��B
'�y�߂�l�z
'  �z�� mAry �̍Ō�̒l�����o���ĕԂ��܂��B
'  mAry ���� (�܂��́A�z��łȂ�) �̏ꍇ�A NULL ���Ԃ���܂��B
'�y�����z
'  �Earray  �̍Ō�̒l�����o���ĕԂ��܂��B
'  �E�z�� array  �́A�v�f����Z���Ȃ�܂��B
'=======================================================================
Function array_pop(ByRef mAry)

    If Not isArray(mAry) Then
        array_pop = null
        Exit Function
    End If

    Dim intCounter
    intCounter = uBound( mAry )
    array_pop = mAry( intCounter )
    ReDim Preserve mAry(intCounter - 1)

End Function