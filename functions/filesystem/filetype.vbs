'=======================================================================
' �t�@�C���^�C�v���擾����
'=======================================================================
'�y�����z
'  filename = string   �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C���̃^�C�v��Ԃ��܂��B
'�y�����z
'  �w�肵���t�@�C���̃^�C�v��Ԃ��܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public Function filetype(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filetype = f.Type

End Function