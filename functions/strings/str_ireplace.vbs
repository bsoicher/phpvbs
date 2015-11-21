'=======================================================================
' �啶������������ʂ��Ȃ� str_replace()
'=======================================================================
'�y�����z
'  search    = mixed ����������
'  strReplace= mixed �u��������
'  subject   = mixed subject  ���z��̏ꍇ�́A���̂��ׂĂ̗v�f�� �΂��Č����ƒu�����s���A�Ԃ���錋�ʂ��z��ƂȂ�܂��B
'  cnt       = mixed needles  �̒��ŁA�}�b�`���Ēu�����s�������� count  �ɕԂ��܂��B���̃p�����[�^�͎Q�Ɠn���Ƃ��܂��B
'�y�߂�l�z
'  �u�����������񂠂邢�͔z���Ԃ��܂��B
'�y�����z
'  �E���̊֐��́Asubject  �̒��Ɍ���邷�ׂĂ� search (�啶������������ʂ��Ȃ�)�� replace  �ɒu�������������񂠂邢�͔z��� �Ԃ��܂��B
'=======================================================================
Function str_ireplace(search, strReplace, subject, ByRef cnt)

    If is_string(search) and isArray(strReplace) Then Exit Function

    If Not isArray(search) Then search = array(search)
    search = array_values(search)

    Dim replace_string,i
    If not isArray(strReplace) Then
        replace_string = strReplace

        strReplace = array()
        For i = 0 to uBound(search)
            [] strReplace, replace_string
        Next
    End if

    strReplace = array_values(strReplace)

    Dim length_replace,length_search
    length_replace = count(strReplace,"")
    length_search  = count(search,"")
    if length_replace < length_search Then
        For i = length_replace to length_search
            strReplace(i) = ""
        Next
    End If

    Dim was_array : was_array = false
    If isArray(subject) Then
        was_array = true
        subject = array( subject )
    End If

    For i = 0 to uBound( search )
        search(i) = "/" & preg_quote(search(i),"") & "/"
    Next

    For i = 0 to uBound( strReplace )
        strReplace(i) = str_replace( array(chr(92),"$"),array(chr(92) & chr(92), "��$"),strReplace(i) )
    Next

    Dim result
    result = preg_replace(search,strReplace,subject,"",cnt)

    If was_array = true Then
        str_ireplace = result(0)
    Else
        str_ireplace = result
    End If

End Function