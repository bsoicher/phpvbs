'=======================================================================
' �t�@�C�����R�s�[����
'=======================================================================
'�y�����z
'  source  = string   �R�s�[���t�@�C���ւ̃p�X�B
'  dest    = string   �R�s�[��̃p�X�B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E �t�@�C�� source  �� dest  �ɃR�s�[���܂��B
'  �E �t�@�C�����ړ��������Ȃ�́Arename() �֐����g�p���Ă��������B 
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================
Public Function copy(source,dest)
    fso.CopyFile source,dest
End Function