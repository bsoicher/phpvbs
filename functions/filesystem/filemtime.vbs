'=======================================================================
' �t�@�C���̍X�V�������擾����
'=======================================================================
'�y�����z
'  filename = string   �t�@�C���ւ̃p�X�B
'�y�߂�l�z
'  �t�@�C���̍ŏI�X�V������Ԃ��A�G���[�̏ꍇ�� FALSE  ��Ԃ��܂��B
'�y�����z
'  ���̊֐��́A�t�@�C���̃u���b�N�f�[�^���������܂ꂽ���Ԃ�Ԃ��܂��B ����́A�t�@�C���̓��e���ύX���ꂽ�ۂ̎��Ԃł��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function filemtime(filename)

    Dim f
    filename = fileMapPath(filename)
    set f = fso.GetFile(filename)
    filemtime = f.DateLastModified

End Function

'*************************************
Private Function fileMapPath(filename)

    Dim tmp
    tmp = Left(filename,3)
    tmp = Lcase(tmp)
    If tmp <> "d:��" and tmp <> "c:��" and left(filename,7) <> "http://" then
            fileMapPath = Server.MapPath(filename)
    Else
        fileMapPath = filename
    End If

End Function