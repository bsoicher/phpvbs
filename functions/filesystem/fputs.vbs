'=======================================================================
' fwrite() �̃G�C���A�X
'=======================================================================
'�y�����z
'  str  =  string �������ޕ�����B
'�y�����z
'  ���̊֐��͎��̊֐��̃G�C���A�X�ł��B fwrite().
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public function fputs(str)
    fputs = fwrite(str)
end function