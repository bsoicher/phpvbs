'=======================================================================
'�����̑������̔z����\�[�g����
'=======================================================================
'�y�����z
'  arr  = array  �\�[�g�������z��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  array_multisort() �́A�������̔z������̎����̈�Ń\�[�g����ۂɎg�p�\�ł��B
'=======================================================================
Function array_multisort(ByRef arr)

    array_multisort = false
    If not isArray(arr) Then Exit Function

    Dim key
    For key = 0 to uBound(arr)
        array_multisort arr(key)
    Next

    sort arr,0

    array_multisort = true

End Function