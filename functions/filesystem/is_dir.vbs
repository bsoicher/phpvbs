'=======================================================================
' �t�@�C�����f�B���N�g�����ǂ����𒲂ׂ�
'=======================================================================
'�y�����z
'  filename = string    �t�@�C���ւ̃p�X�Bfilename  �����΃p�X�̏ꍇ�́A���݂̍�ƃf�B���N�g������̑��΃p�X�Ƃ��ď������܂��B
'�y�߂�l�z
'  �t�@�C���������݂��āA�����ꂪ�f�B���N�g���ł���� TRUE�A����ȊO�̏ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �w�肵���t�@�C�����f�B���N�g�����ǂ����𒲂ׂ܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public Function is_dir(filename)

    Dim fn
    is_dir = false
    fn = fileMapPath(filename)

    If fso.FolderExists(fn) Then is_dir = true

End Function