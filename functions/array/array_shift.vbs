'=======================================================================
'�z��̐擪����v�f������o��
'=======================================================================
'�y�����z
'  ary     = Array �z��
'�y�߂�l�z
'  �Eary  �̍ŏ��̒l�����o���ĕԂ��܂��B
'  �Earray  ����̏ꍇ (�܂��͔z��łȂ��ꍇ)�A NULL ���Ԃ���܂��B
'�y�����z
'  �E�z�� ary  �́A�v�f��������Z���Ȃ�A�S�Ă̗v�f�͑O�ɂ���܂��B 
'  �E���l�Y���̔z��̃L�[�̓[�����珇�ɐV���ɐU��Ȃ�����܂����A ���e�����̃L�[�͂��̂܂܂ɂȂ�܂��B
'=======================================================================
Function array_shift(ByRef ary)

    If Not isArray(ary) and Not isObject(ary) then
        array_shift = null
        Exit Function
    End If

    Dim i,key : i = 0

    If isArray(ary) Then
        array_shift = ary(0)

        For i = 0 to uBound(ary)-1
            ary(i) = ary(i+1)
        Next
        Redim Preserve ary(UBound(ary) - 1)

    ElseIf isObject(ary) Then
        For Each key In ary
            array_shift = ary(key)
            ary.Remove(key)
            Exit For
        Next
    End if

End Function