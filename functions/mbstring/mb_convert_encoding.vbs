'=======================================================================
' �����G���R�[�f�B���O��ϊ�����
'=======================================================================
'�y�����z
'  str          = string    �ϊ����镶����B
'  to_encoding  = string    str  �̕ϊ���̕����G���R�[�f�B���O�B
'  from_encoding= string    �ϊ��O�̕����G���R�[�f�B���O�����w�肵�܂��B
'�y�߂�l�z
'  �ϊ���̕������Ԃ��܂��B
'�y�����z
'  �E������ str �̕����G���R�[�f�B���O���A �I�v�V�����Ŏw�肵�� from_encoding  ���� to_encoding  �ɕϊ����܂��B
'=======================================================================
Function mb_convert_encoding(str,to_encoding,from_encoding)

    Dim bobj : set bobj = Server.CreateObject("basp21")
    mb_convert_encoding = bobj.Kconv(str,_
                          mb_convert_encoding_helper(to_encoding),_
                          mb_convert_encoding_helper(from_encoding))
End Function

'*******************************************
Function mb_convert_encoding_helper(encoding)

    Dim tmp
    Select Case lcase(encoding)
    Case "shift_jis","sjis"
        tmp = 1
    Case "euc","euc-jp"
        tmp = 2
    Case "jis"
        tmp = 3
    Case "ucs2"
        tmp = 4
    Case "utf-8","utf8"
        tmp = 5
    Case "auto"
        tmp = 0
    End Select

End Function