'=======================================================================
'�z��̂��ׂẴL�[��ύX����
'=======================================================================
'�y�����z
'  mObj     = objec  �������s���A�z�z��B
'  flag     = int    CASE_UPPER ���邢�� CASE_LOWER (�f�t�H���g)�B
'�y�߂�l�z
'  ���ׂẴL�[�����������邢�͑啶���ɂ����z���Ԃ��܂��B input  ���z��łȂ��ꍇ�� false ��Ԃ��܂��B
'�y�����z
'  �EmAry  �̂��ׂẴL�[�����������邢�͑啶���ɂ����z���Ԃ��܂��B ���l�Y���͂��̂܂܂ƂȂ�܂��B
'=======================================================================
Const CASE_UPPER = 1
Const CASE_LOWER = 0
Function array_change_key_case(ByRef mObj, flag)

    if flag <> CASE_UPPER and flag <> CASE_LOWER then Exit Function
    if Not isObject(mObj) Then Exit Function

    Dim cnt,i

    arykey = mObj.Keys
    aryval = mObj.Items
    cnt    = mObj.Count -1
    mObj.RemoveAll

    Select Case flag
    Case CASE_UPPER

        For i = 0 to cnt
            mObj.Add Ucase(arykey(i)), aryval(i)
        Next

    Case CASE_LOWER

        For i = 0 to cnt
            mObj.Add Lcase(arykey(i)), aryval(i)
        Next

    End Select

End Function