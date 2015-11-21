'=======================================================================
' ������� rot13 �ϊ����s��
'=======================================================================
'�y�����z
'  str  = string    ���͕�����B
'�y�߂�l�z
'  �w�肵��������� ROT13 �ϊ��������ʂ�Ԃ��܂��B
'�y�����z
'  �EROT13 �́A�e�������A���t�@�x�b�g���� 13 �����V�t�g�����A �A���t�@�x�b�g�ȊO�̕����͂��̂܂܂Ƃ���G���R�[�h���s���܂��B
'  �G���R�[�h�ƃf�R�[�h�͓����֐��ōs���܂��B
'  �����ɃG���R�[�h���ꂽ��������w�肵���ꍇ�ɂ́A���̕����񂪕Ԃ���܂��B
'=======================================================================
Function str_rot13(str)

    Dim str_rotated : str_rotated = ""
    Dim i,j,k

    For i = 1 to Len(str)
        j = Mid(str, i, 1)
        k = Asc(j)
        if k >= 97 and k =< 109 then
            k = k + 13 ' a ... m
        elseif k >= 110 and k =< 122 then
            k = k - 13 ' n ... z
        elseif k >= 65 and k =< 77 then
            k = k + 13 ' A ... M
        elseif k >= 78 and k =< 90 then
            k = k - 13 ' N ... Z
        end if

        str_rotated = str_rotated & Chr(k)
    Next

    str_rot13 = str_rotated

End Function