'=======================================================================
' �t�@�C���̃T�C�Y���擾����
'=======================================================================
'�y�����z
'  filename = string   �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C���̃T�C�Y��Ԃ��A�G���[�̏ꍇ�� FALSE ��Ԃ��܂� (�܂� E_WARNING ���x���̃G���[�𔭐������܂�) �B
'�y�����z
'  �w�肵���t�@�C���̃T�C�Y���擾���܂��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function filesize(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filesize = f.Size

End Function