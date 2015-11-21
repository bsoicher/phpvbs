'=======================================================================
' "���R��"�A���S���Y���ɂ�蕶�����r���s��
'=======================================================================
'�y�����z
'  str1 =   string  �ŏ��̕�����B
'  str2 =   string  ���̕�����B
'�y�߂�l�z
'  ���̕������r�֐��Ɠ��l�ɁA���̊֐��́A str1  �� str2  ��菬�����ɏꍇ�� < 0�Astr1  �� str2  ���傫���ꍇ�� > 0 �A�������ꍇ�� 0 ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A�l�Ԃ��s���悤�Ȏ�@�ŃA���t�@�x�b�g�܂��͐����� ������̏������r����A���S���Y�����������܂��B���̎�@�́A"���R��" �ƌ����܂��B
'  �E���̔�r�́A�啶������������ʂ��邱�Ƃɒ��ӂ��Ă��������B
'=======================================================================
Function strnatcmp( str1, str2 )

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
    strnatcmp = false
    For i = 0 to intlen
        If not isNumeric( array1(i) ) Then
            If Not isNumeric( array2(i) ) Then
                text = true

                r = strcmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcmp = r
                End If

            ElseIf text Then
                strnatcmp = 1
            Else
                strnatcmp = 1
            End If

        ElseIf not isNumeric( array2(i) ) Then
            If text Then
                strnatcmp = -1
            Else
                strnatcmp = 1
            End If
        Else
            If text Then
                r = array1(i) - array2(i)
                If r <> 0 Then
                    strnatcmp = r
                End If
            Else
                r = strcmp(Cstr( array1(i) ),Cstr( array2(i) ) )
                If r <> 0 Then
                    strnatcmp = r
                End If
            End If

            text = false
        End If

        if [!==](strnatcmp,false) Then Exit Function
    Next

    strnatcmp = result

End Function

'*****************************
Function strnatcmp_helper(str)

    Dim result
    Dim buffer : buffer = ""
    Dim strChr : strChr = ""
    Dim text   : text   = true

    Dim i
    For i = 1 to len(str)
        strChr = Mid(str,i,1)

        If preg_match("/[0-9]/i",strChr,"","","") Then
            If text Then
                If len( buffer ) > 0 Then
                    [] result, buffer
                    buffer = ""
                End If

                text = false
            End If
            buffer  = buffer & strChr
        ElseIf text = false and strChr = "." and i < (len( str )-1) Then
            If preg_match("/[0-9]/",Mid(str,i+1,1),"","","") Then
                [] result, buffer
                buffer = ""
            End If
        Else
            If text = false Then
                If len( buffer ) > 0 Then
                    [] result , intval(buffer)
                    buffer = ""
                End If
                text = true
            End If
            buffer = buffer & strChr
        End If
    Next

    If len( buffer ) > 0 Then
        if text Then
            [] result, buffer
        Else
            [] result, intval(buffer)
        End If
    End If

    strnatcmp_helper = result
End Function