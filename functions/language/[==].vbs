'=======================================================================
'a �� b �ɓ��������� TRUE�B
'=======================================================================
'�y�����z
'  a    = mixed  �l
'  b    = mixed  ��r����l
'�y�߂�l�z
'  a��b���������ꍇ��TRUE ���A�������Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���ӂƉE�ӂ��r���܂��B�^�͌����Ƀ`�F�b�N���܂���B
'=======================================================================
Function [==](ByVal a, ByVal b)

    [==] = false

    Dim tmp_a,tmp_b
    Dim key

    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count <> b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [==](tmp_a,tmp_b) Then Exit Function

            tmp_a = a.Items : tmp_b = b.Items
            If Not [==](tmp_a,tmp_b) Then Exit Function
            [==] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) <> uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [==](a(key),b(key) ) Then Exit Function
            Next

            [==] = true
        End If

    Else
        If isNull(a) Then a = ""
        If isNull(b) Then b = ""

        [==] = (Cstr(a) = Cstr(b))
    End If

End Function