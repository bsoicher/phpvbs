'=======================================================================
'�z��ɒl�����邩�`�F�b�N����
'=======================================================================
'�y�����z
'  needle     = mixed �T���l�B
'  haystack   = Array  �z��B
'  strict     = bool   �O�Ԗڂ̃p�����[�^ strict �� TRUE �ɐݒ肳�ꂽ�ꍇ�A haystack �̒��� needle �̌^���m�F���܂��B
'�y�߂�l�z
'  �z��� needle  �����������ꍇ�� TRUE�A����ȊO�̏ꍇ�́AFALSE ��Ԃ��܂��B
'�y�����z
'  �Ehaystack�z�����needle���܂܂�邩�`�F�b�N
'=======================================================================
Function in_array(needle, haystack,strict)

    in_array = False

    If Not IsArray(needle) Then
        If Len( needle ) = 0 Then Exit Function
    End If

    If VarType(strict) <> 11 Then strict = false

    Dim key

    If isArray(needle) Then
        For Each key In needle
            in_array = in_array(key,haystack,strict)
            If in_array = true Then Exit For
        Next
        Exit Function
    End If

    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    End If

End Function