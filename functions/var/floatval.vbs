'=======================================================================
'�ϐ��� float �l���擾����
'=======================================================================
'�y�����z
'  str  = mixed  ������X�J���^���w��ł��܂��B�z�񂠂邢�̓I�u�W�F�N�g�� floatval() ���g�p���邱�Ƃ͂ł��܂���B
'�y�߂�l�z
'  �w�肵���ϐ��� float �l��Ԃ��܂��B
'�y�����z
'  �E�ϐ� str �� float �l��Ԃ��܂��B
'=======================================================================
Function floatval(str)

    floatval = false
    If isArray(str) or isObject(str) Then Exit Function
    If not isNumeric(str) Then Exit Function
    floatval = CDbl(str)

End Function