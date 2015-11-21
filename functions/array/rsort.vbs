'=======================================================================
'�z����t���Ƀ\�[�g����
'=======================================================================
'�y�����z
'  ary        = Array   ���͂̔z��B
'  sort_flags = int     �I�v�V�����̃p�����[�^ sort_flags  �ɂ��\�[�g�̓�����C���\�ł��B�ڍׂɂ��ẮA sort() ���Q�Ƃ��������B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A�z����t����(���ʂ����ʂ�)�\�[�g���܂��B
'=======================================================================
Function rsort(ByRef arr, sort_flags)

    rsort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While rsort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    rsort = true

End Function

'******************************************
Function rsort_helper(temp , arr, sort_flags)

    rsort_helper = true
    If isArray(temp) or isObject(temp) Then Exit Function

    rsort_helper = false
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        rsort_helper = (temp > arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        rsort_helper = (intval(temp) > intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        rsort_helper = (Cstr(temp) > Cstr(arr))
    End If

End Function