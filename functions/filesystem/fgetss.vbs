'=======================================================================
' �t�@�C���|�C���^���� 1 �s���o���AHTML �^�O����菜��
'=======================================================================
'�y�����z
'  
'�y�߂�l�z
'  HTML �� PHP �R�[�h����菜�����������Ԃ��܂��B
'�y�����z
'  fgets() �Ɠ����ł����A fgetss() �͓ǂݍ��񂾃e�L�X�g���� HTML ����� PHP �̃^�O����菜�����Ƃ��邱�Ƃ��قȂ�܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public function fgetss
    fgets = strip_tags(ts.ReadLine)
end function