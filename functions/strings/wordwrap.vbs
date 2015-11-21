'=======================================================================
' �����񕪊��������g�p���Ďw�肵�����������ɕ�����𕪊�����
'=======================================================================
'�y�����z
'  str      = string    ���͕�����B
'  width    = int       �J�����̕��B�f�t�H���g�� 75�B
'  break    = string    �I�v�V�����̃p�����[�^ break  ��p���čs�𕪊����܂��B �f�t�H���g�� 'vbCrLf' �ł��B
'  cut      = bool      cut  �� TRUE �ɐݒ肷��ƁA ������͏�Ɏw�肵�����Ń��b�v����܂��B���̂��߁A �w�肵�������������P�ꂪ����ꍇ�ɂ́A��������܂� (2 �Ԗڂ̗���Q�Ƃ�������)�B
'�y�߂�l�z
'  �ϊ���̕������Ԃ��܂��B
'�y�����z
'  �E �w�肵���������ŁA�w�肵��������p���ĕ�����𕪊����܂��B
'=======================================================================
Function wordwrap( str, int_width, str_break, cut )

    If len(int_width) = 0 Then int_width = 75
    If len(str_break) = 0 Then str_break = vbCrLf

    Dim m : m = int_width
    Dim b : b = str_break
    Dim c : c = cut

    Dim i,j, l, s, r
    Dim matches

    If m < 1 Then
        wordwrap = str
        Exit Function
    End If

    r = split(str,vbCrLf)
    l = uBound(r)
    i = -1

    Do While i < l
        i = i +1

        s = r(i)
        r(i) = ""

        Do While len(s) > m
            j = [==](c, 2)
            If is_empty(j) Then
                If preg_match("/��S*(��s)?$/",Left(s,m+1),matches,"","") Then
                    If len( trim(matches(0)) ) = 0 Then
                        j = m
                    Else
                        j = len( Left(s,m+1) ) - len(matches(0))
                    End If
                End If

                If is_empty(j) Then
                    j = [?]([==](c, true),m,false)
                End If

                If is_empty(j) Then
                    call preg_match("/^��S*/",Mid(s,m),matches,"","")
                    j = len( Left(s,m) ) + len(matches(0))
                End If
            End If

            r(i) = r(i) & Left(s, j)
            s = Mid(s,j+1)
            r(i) = r(i) & [?](len(s), b , "")
        Loop

        r(i) = r(i) & s

    Loop

    wordwrap = join(r,vbCrLf)

End Function