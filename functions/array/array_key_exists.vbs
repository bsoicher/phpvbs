'=======================================================================
'�w�肵���L�[�܂��͓Y�����z��ɂ��邩�ǂ����𒲂ׂ�
'=======================================================================
'�y�����z
'  key      = mixed  �z��
'  sAry     = array  �L�[�����݂��邩�ǂ����𒲂ׂ����z��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E�w�肵�� key  ���z��ɐݒ肳��Ă���ꍇ�A array_key_exists() �� TRUE ��Ԃ��܂��B 
'  �Ekey  �͔z��Y���Ƃ��Ďg�p�ł���S�Ă̒l���g�p�\�ł��B
'=======================================================================
Function array_key_exists(key, sAry)

    array_key_exists = false
    If isObject(sAry) Then
        if sAry.Exists( key ) then array_key_exists = true
    ElseIf isArray(sAry) and isNumeric(key) Then
        If (uBound(sAry)+1) > key Then
            If Not isNull(sAry(key)) Then array_key_exists = true
        End If
    End If

End Function