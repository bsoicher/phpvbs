'=======================================================================
' �t�@�C���܂��̓f�B���N�g�������݂��邩�ǂ������ׂ�
'=======================================================================
'�y�����z
'  path      = string   �t�@�C�����邢�̓f�B���N�g���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C�����邢�̓f�B���N�g�������݂��邩�ǂ����𒲂ׂ܂��B
'�y�����z
'  �t�@�C�����邢�̓f�B���N�g�������݂��邩�ǂ����𒲂ׂ܂��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function file_exists(ByVal filename)

    file_exists = false
    filename = fileMapPath(filename)
    If fso.FileExists(filename) Then file_exists = true
    If fso.FolderExists(filename) Then file_exists = true

End Function
