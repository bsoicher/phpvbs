'=======================================================================
'�z����\�[�g����
'=======================================================================
'�y�����z
'  ary        = Array   �\�[�g����z��
'  sort_flags = int     �\�[�g�̓�����C�����邽�߂Ɏg�p���邱�Ƃ��\�ł��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E�z����\�[�g���܂��B
'�@�E���̊֐�������ɏI������ƁA �e�v�f�͒�ʂ��獂�ʂ֕��בւ����܂��B
'�@�Ehttp://www.thinkit.co.jp/article/62/3/
'=======================================================================

Const SORT_REGULAR = 0
Const SORT_NUMERIC = 1
Const SORT_STRING  = 2

Function sort(ByRef arr, sort_flags)

    sort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While sort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    sort = true

End Function

'******************************************
Function sort_helper(temp , arr, sort_flags)

    sort_helper = false
    If isArray(temp) or isObject(temp) Then Exit Function

    sort_helper = true
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        sort_helper = (temp < arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        sort_helper = (intval(temp) < intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        sort_helper = (Cstr(temp) < Cstr(arr))
    End If

End Function

