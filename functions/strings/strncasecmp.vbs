'=======================================================================
' �o�C�i���Z�[�t�ő啶������������ʂ��Ȃ��������r���A�ŏ��� n �����ɂ��čs��
'=======================================================================
'�y�����z
'  str1     =   string  �ŏ��̕�����B
'  str2     =   string  ���̕�����B
'  intlen   =   string  ��r���镶����̒����B
'�y�߂�l�z
' str1  �� str2  ���Z���ꍇ�� < 0 ��Ԃ��Astr1  �� str2  ���傫���ꍇ�� > 0�A�������ꍇ�� 0 ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́Astrcasecmp() �Ɏ��Ă��܂����A �e�����񂩂��r���镶����(�̏��)(len ) ���w��ł���Ƃ����Ⴂ������܂��B
'  �E�ǂ��炩�̕����� len ���Z���ꍇ�A���̕�����̒�������r���Ɏg�p����܂��B
'=======================================================================
Function strncasecmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncasecmp = StrComp(str1,str2,vbTextCompare)
End Function