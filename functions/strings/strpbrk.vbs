'=======================================================================
' ������̒�����C�ӂ̕�����T��
'=======================================================================
'�y�����z
'  haystack     =   string  char_list  ��T��������B
'  char_list    =   string  ���̃p�����[�^�͑啶������������ʂ��܂��B
'�y�߂�l�z
'  ����������������n�܂镶����A���邢�͌�����Ȃ������ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E strpbrk() �́A������ haystack  ���� char_list  ��T���܂��B
'=======================================================================
Function strpbrk( haystack, char_list )

    haystack  = Cstr( haystack )
    char_list = Cstr( char_list )

    Dim intlen : intlen = len( haystack )
    Dim i,char
    For i = 1 to intlen
        char = Mid(haystack,i,1)
        If [!==](strpos(char_list,char,""),false) Then
            strpbrk = Mid(haystack,i)
            Exit Function
        End If
    Next

    strpbrk = false

End Function