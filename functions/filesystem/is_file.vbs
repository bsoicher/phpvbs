'=======================================================================
' �ʏ�t�@�C�����ǂ����𒲂ׂ�
'=======================================================================
'�y�����z
'  filename = string    �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C�������݂��A�����ꂪ�ʏ�̃t�@�C���ł���ꍇ�� TRUE�A ����ȊO�̏ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �w�肵���t�@�C�����ʏ�̃t�@�C�����ǂ����𒲂ׂ܂��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function is_file(ByVal filename)

    is_file = false
    filename = fileMapPath(filename)
    If fso.FileExists(filename) Then is_file = true

End Function