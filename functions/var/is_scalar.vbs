'=======================================================================
'�ϐ����X�J�����ǂ����𒲂ׂ�
'=======================================================================
'�y�����z
'  str   = mixed �]������ϐ�
'�y�߂�l�z
'  str ���X�J���̏ꍇ TRUE�A �����łȂ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E �w�肵���ϐ����X�J�����ǂ����𒲂ׂ܂��B
'  �E �X�J���ϐ��ɂ� integer�Afloat�Astring ���邢�� boolean ���܂܂�܂��B
'  �E array�Aobject ����� resource �̓X�J���ł͂���܂���B 
'=======================================================================
Function is_scalar(str)

    is_scalar = false
    If isArray(str) or isObject(str) Then Exit Function
    if isNull(str) Then Exit Function
    is_scalar = true

End Function