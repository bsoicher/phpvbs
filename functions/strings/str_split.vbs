'=======================================================================
' �������z��ɕϊ�����
'=======================================================================
'�y�����z
'  string       = string ���͕�����B
'  split_length = string �������������̍ő咷�B
'�y�߂�l�z
'  �I�v�V�����̃p�����[�^ split_length  ���w�肳��Ă���ꍇ�A �Ԃ����z��̊e�v�f�́Asplit_length  �̒����ƂȂ�܂��B����ȊO�̏ꍇ�A1 �������������ꂽ�z��ƂȂ�܂��B
'  split_length �� 1 ��菬�����ꍇ�� FALSE ��Ԃ��܂��B
'  split_length �� string �̒������傫���ꍇ�A������S�̂� �ŏ���(�����ėB���)�v�f�ƂȂ�z���Ԃ��܂��B 
'�y�����z
'  �E�������z��ɕϊ����܂��B
'=======================================================================
Function str_split(string, split_length)

    str_split = false
    If len(string) = 0 Then Exit Function
    If len(split_length) = 0 Then split_length = 1
    If split_length < 1 Then Exit Function

    Dim counter,i,pointer
    counter = len(string)
    counter = counter / split_length + 0.9999
    counter = int(counter) -1

    ReDim tmp_ar(counter)

    For i = 0 to counter
        pointer = i * split_length + 1
        tmp_ar(i) = Mid(string,pointer,split_length)
    Next

    str_split = tmp_ar

End Function