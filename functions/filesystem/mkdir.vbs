'=======================================================================
' �f�B���N�g�������
'=======================================================================
'�y�����z
'  pathname = string    �f�B���N�g���̃p�X�B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �w�肵���f�B���N�g�����쐬���܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public Function mkdir(ByVal pathname)

    mkdir = false
    pathname = fileMapPath(pathname)
    If not fso.FolderExists(pathname) Then
        mkdir = fso.CreateFolder(pathname)
    End If

End Function