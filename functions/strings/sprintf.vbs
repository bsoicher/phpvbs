'=======================================================================
' �t�H�[�}�b�g���ꂽ�������Ԃ�
'=======================================================================
'�y�����z
'  format   = string �t�H�[�}�b�g������
'  args     = mixed  ���l�╶����
'�y�߂�l�z
'  �t�H�[�}�b�g������ format  �Ɋ�Â��������ꂽ�������Ԃ��܂��B
'�y�����z
'  �E�t�H�[�}�b�g������ format  �Ɋ�Â��������ꂽ�������Ԃ��܂��B
'=======================================================================
Function sprintf(format , args)

    If is_empty(args) Then args = ""
    Dim bobj : set bobj = Server.CreateObject("basp21")
    sprintf = bobj.sprintf(format,args)

End Function