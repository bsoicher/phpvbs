'=======================================================================
'a �� b �ɓ������Ȃ����� TRUE�B
'=======================================================================
'�y�����z
'  a    = mixed  �l
'  b    = mixed  ��r����l
'�y�߂�l�z
'  a��b���������Ȃ��ꍇ��TRUE ���A�������Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���ӂƉE�ӂ��r���܂��B�^�͌����Ƀ`�F�b�N���܂���B
'=======================================================================
Function [!=](a, b)

    If [==](a,b) Then
        [!=] = false
    Else
        [!=] = true
    End If

End Function