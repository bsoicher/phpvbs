'=======================================================================
' �ŏ��� n �����ɂ��ăo�C�i���Z�[�t�ȕ������r���s��
'=======================================================================
'�y�����z
'  str1     =   string  �ŏ��̕�����B
'  str2     =   string  ���̕�����B
'  intlen   =   string  ��r���镶�����B
'�y�߂�l�z
' str1  �� str2  ���Z���ꍇ�� < 0 ��Ԃ��Astr1  �� str2  ���傫���ꍇ�� > 0�A�������ꍇ�� 0 ��Ԃ��܂��B
'�y�����z
'  �E ���̊֐��� strcmp() �Ɏ��Ă��܂����A �e�����񂩂�(�ő�)������(len ) ���r�Ɏg�p����Ƃ��낪�قȂ�܂��B
'  �E ��r�͑啶������������ʂ��邱�Ƃɒ��ӂ��Ă��������B 
'=======================================================================
Function strncmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncmp = StrComp(str1,str2,vbBinaryCompare)
End Function