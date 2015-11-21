'=======================================================================
'����������Ɉ�v�������ׂĂ̕������u������
'=======================================================================
'�y�����z
'  search    = mixed ����������
'  replace   = mixed �u��������
'  subject   = mixed �u���Ώە�����
'�y�߂�l�z
'  �u����̕����񂠂邢�͔z���Ԃ��܂��B
'�y�����z
'  �Esubject  �̒��� search  ��S�� replace  �ɒu�����܂��B
'=======================================================================
Function str_replace(ByVal search, ByVal strReplace, ByVal subject)

    Dim tmp
    Dim J

    If IsObject(search) or IsObject(strReplace) or IsObject(subject) Then Exit Function

    If IsArray(search) and Not IsArray(strReplace) Then
        tmp = strReplace
        ReDim strReplace(UBound(search))
        strReplace(0) = tmp
    ElseIf Not IsArray(search) and IsArray(strReplace) Then
        tmp = search
        ReDim search(UBound(strReplace))
        search(0) = tmp
    End If

    If IsArray(search) and IsArray(strReplace) Then

        If UBound(search) <> UBound(strReplace) Then

            If UBound(search) > UBound(strReplace) Then

                ReDim strReplace(UBound(search))

            ElseIf UBound(search) < UBound(strReplace) Then

                ReDim search(UBound(strReplace))

            End If

        End If

    End If

    If IsArray(subject) Then
        For J = 0 To UBound(subject)
            subject(J) = str_replace(search, strReplace, subject(J))
        Next

    Else

        If IsArray(search) Then
            For J = 0 To UBound(search)
                subject = Replace(subject,search(J),strReplace(J),1,len(subject),vbBinaryCompare)
            Next
        Else
            subject = Replace(subject,search,strReplace,1,len(subject),vbBinaryCompare)
        End If

    End If

    str_replace = subject

End Function