'=======================================================================
' �f�B���N�g�����폜����
'=======================================================================
'�y�����z
'  dirname     = string    �f�B���N�g���ւ̃p�X�B
'�y�߂�l�z
'   ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  dirname �Ŏw�肳�ꂽ�f�B���N�g���� �폜���悤�Ǝ��݂܂��B
'  �f�B���N�g���͋�łȂ��Ă͂Ȃ炸�A�܂� �K�؂ȃp�[�~�b�V�������ݒ肳��Ă��Ȃ���΂Ȃ�܂���B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function rmdir(ByVal dirname)

    dirname = fileMapPath(dirname)
    fso.DeleteFolder dirname
    rmdir = true

End Function