'=======================================================================
' �t�@�C�����폜����
'=======================================================================
'�y�����z
'  filename     = string    �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'   ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  filename  ���폜���܂��B Unix C ����̊֐� unlink() �Ɠ���͓����ł��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public Function unlink(ByVal filename)

    filename = fileMapPath(filename)
    fso.DeleteFile filename
    unlink = true

End Function