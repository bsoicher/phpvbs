'=======================================================================
' �����񒆂ɕ������Ō�Ɍ����ꏊ���擾����
'=======================================================================
'�y�����z
'  haystack = string    �������s��������B
'  needle   = string    needle  ���ЂƂȏ�̕������܂�ł���ꍇ�́A �ŏ��̂��݂̂̂��g���܂��B���̓���́A strstr()  �Ƃ͈قȂ�܂��B
'�y�߂�l�z
'  ���̊֐��́A�����������Ԃ��܂��B needle  ��������Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E ���̊֐��́A������ haystack  �̒��� needle  ���Ō�Ɍ��ꂽ�ʒu����A haystack  �̏I���܂ł�Ԃ��܂��B
'=======================================================================
Function strrchr( haystack, needle )

    haystack = Cstr( haystack )
    needle   = Cstr( needle )
    If len(needle) > 1 Then needle = Left(needle,1)

    strrchr = false

    Dim i
    i = strrpos(haystack, needle,"")

    If i > 0 Then
        strrchr = Mid(haystack,i)
    End If

End Function