'=======================================================================
' �t�@�C���̍ŏI�A�N�Z�X�������擾����
'=======================================================================
'�y�����z
'  filename = string   �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C���̍ŏI�A�N�Z�X������Ԃ��A�G���[�̏ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �w�肵���t�@�C���̍ŏI�A�N�Z�X�������擾���܂��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function fileatime(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    fileatime = f.DateLastAccessed

End Function