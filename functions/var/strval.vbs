'=======================================================================
'�ϐ��̕�����Ƃ��Ă̒l�𓾂܂�
'=======================================================================
'�y�����z
'  str   = mixed ������ɕϊ��������ϐ��Bstr �́A�S�ẴX�J���[�l�ɂł��܂��B strval() �ɔz�񂠂邢�̓I�u�W�F�N�g�͎g�p�ł��܂���B
'�y�߂�l�z
'  str������l��Ԃ��܂��B
'�y�����z
'  �Estr��string �Ƃ��Ă̒l��Ԃ��܂��B
'=======================================================================
Function strval(ByVal str)

    strval = false
    If isArray(str) or isObject(str) Then Exit Function
    strval = Cstr(str)

End Function