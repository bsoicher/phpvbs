'=======================================================================
' "���R��"�A���S���Y���ɂ��啶������������ʂ��Ȃ��������r���s��
'=======================================================================
'�y�����z
'  str1 =   string  �ŏ��̕�����B
'  str2 =   string  ���̕�����B
'�y�߂�l�z
'  ���̕������r�֐��Ɠ��l�ɁA���̊֐��́A str1  �� str2  ��菬�����ɏꍇ�� < 0�Astr1  �� str2  ���傫���ꍇ�� > 0 �A�������ꍇ�� 0 ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A�l�Ԃ��s���悤�Ȏ�@�ŃA���t�@�x�b�g�܂��͐����� ������̏������r����A���S���Y�����������܂��B���̎�@�́A"���R��" �ƌ����܂��B
'  �E���̊֐��̓���́A strnatcmp() �Ɏ��Ă��܂����A ��r���啶������������ʂ��Ȃ��Ⴂ������܂��B
'=======================================================================
Function strnatcasecmp( str1, str2 )

    Dim array1,array2
    array1 = strnatcmp_helper(str1)
    array2 = strnatcmp_helper(str2)

    Dim intlen,text,result,r
    intlen = uBound(array1)
    text   = true

    result = -1
    r      = 0

    if intlen > uBound(array2) Then
        intlen = uBound(array2)
        result = 1
    End If

    Dim i
    strnatcasecmp = false
    For i = 0 to intlen
        If not isNumeric( array1(i) ) Then
            If Not isNumeric( array2(i) ) Then
                text = true

                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If

            ElseIf text Then
                strnatcasecmp = 1
            Else
                strnatcasecmp = 1
            End If

        ElseIf not isNumeric( array2(i) ) Then
            If text Then
                strnatcasecmp = -1
            Else
                strnatcasecmp = 1
            End If
        Else
            If text Then
                r = array1(i) - array2(i)
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            Else
                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            End If

            text = false
        End If

        if [!==](strnatcasecmp,false) Then Exit Function
    Next

    strnatcasecmp = result

End Function