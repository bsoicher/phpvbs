'=======================================================================
' ��������̏o���񐔂𐔂���
'=======================================================================
'�y�����z
'  haystack     = string    �����Ώۂ̕�����
'  needle       = string    �������镛������
'  offset       = int       �J�n�ʒu�̃I�t�Z�b�g
'  length       = int       �w�肵���I�t�Z�b�g�ȍ~�ɕ�������Ō�������ő咷�B
'�y�߂�l�z
'  ���̊֐��� ���� ��Ԃ��܂��B
'�y�����z
'  �Esubstr_count() �́A������ haystack  �̒��ł̕������� needle  �̏o���񐔂�Ԃ��܂��B
'  �Eneedle  �͉p�召��������ʂ��邱�Ƃɒ��ӂ��Ă��������B
'=======================================================================
Function substr_count( haystack, needle, offset, length )

    Dim pos,cnt : cnt = 0

    If not isNumeric(offset) Then offset = 1
    If not isNumeric(length) Then length = 0

    Do While inStr(offset+1,haystack,needle,vbBinaryCompare) > 0
        offset = inStr(offset+1,haystack,needle,vbBinaryCompare)
        If length > 0 and offset + len(needle) > length Then
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop

    substr_count = cnt

End Function