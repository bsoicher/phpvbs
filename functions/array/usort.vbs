'=======================================================================
'���[�U�[��`�̔�r�֐����g�p���āA�z���l�Ń\�[�g����
'=======================================================================
'�y�����z
'  ary          = Array   ���͂̔z��B
'  cmp_function = int     ��r�֐��́A�ŏ��̈����� 2 �Ԗڂ̈�����菬�������A���������A�傫���ꍇ�ɁA ���ꂼ��[�������A�[���ɓ������A�[�����傫��������Ԃ� �K�v������܂��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A���[�U�[��`�̔�r�֐��ɂ��z������̒l�Ń\�[�g���܂��B 
'  �E�\�[�g�������z��𕡎G�Ȋ�Ń\�[�g����K�v������ꍇ�A ���̊֐����g�p����ׂ��ł��B
'=======================================================================
Function usort(ByRef arr, cmp_function)

    usort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)
        j = i -1
        Do While usort_helper(temp,arr(j),cmp_function)
            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    usort = true


End Function

'*****************************************************
Function usort_helper(temp,arr,cmp_function)

    Dim output
    execute ("output = " & cmp_function & "(temp,arr)")
    usort_helper = output

End Function