'=======================================================================
'�ϐ��̐����Ƃ��Ă̒l���擾����
'=======================================================================
'�y�����z
'  var = mixed ������
'�y�߂�l�z
'  ����
'�y�����z
'  �Evar  �� integer �Ƃ��Ă̒l��Ԃ��܂��B
'=======================================================================
Function intval(str)

    intval = 1
    If IsObject(str) or IsArray(str) Then Exit Function
    If str = true Then Exit Function

    intval = 0
    If is_empty(str) or Not isNumeric(str) Then Exit Function

    str = int(str)
    If str > 32767 Then
        intval = 32767
    Else
        intval = Cint(str)
    End If

End Function