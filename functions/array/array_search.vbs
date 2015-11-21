'=======================================================================
'�w�肵���l��z��Ō������A���������ꍇ�ɑΉ�����L�[��Ԃ�
'=======================================================================
'�y�����z
'  needle   = mixed �T�������l�B
'  haystack = array �z��B
'  strict   = mixed TRUE ���w�肳�ꂽ�ꍇ�Aarray_search() �� haystack  �̒��� needle  �̌^�Ɉ�v���邩�ǂ������m�F���܂��B
'�y�߂�l�z
'  �Eneedle  �����������ꍇ�ɔz��̃L�[�A ����ȊO�̏ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Ehaystack  �ɂ����� needle  ���������܂��B
'=======================================================================
Function array_search(needle,haystack,strict)

    array_search = false

    If IsArray(needle) or isObject(needle) Then Exit Function

    If VarType(strict) <> 11 Then strict = false

    Dim key
    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    End If

End Function