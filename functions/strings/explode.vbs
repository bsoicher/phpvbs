'=======================================================================
' ������𕶎���ɂ�蕪������
'=======================================================================
'�y�����z
'  delimiter    = string    ��؂蕶����B
'  string       = string    ���͕�����B
'  limit        = string    limit  ���w�肳�ꂽ�ꍇ�A�Ԃ����z��ɂ� �ő� limit  �̗v�f���܂܂�A���̍Ō�̗v�f�ɂ� string  �̎c��̕������S�Ċ܂܂�܂��B
'�y�߂�l�z
'  ��̕����� ("") �� delimiter  �Ƃ��Ďg�p���ꂽ�ꍇ�A explode() �� FALSE  ��Ԃ��܂��B
'  delimiter  �Ɉ��� string  �Ɋ܂܂�Ă��Ȃ��l���܂܂�Ă���ꍇ�A explode() �́A���� string  ���܂ޔz���Ԃ��܂��B
'�y�����z
'  �E������̔z���Ԃ��܂��B���̔z��̊e�v�f�́A string  �𕶎��� delimiter  �ŋ�؂�������������ƂȂ�܂��B
'=======================================================================
Function explode(delimiter,string,limit)

    explode = false
    If len(delimiter) = 0 Then Exit Function
    If len(limit) = 0 Then limit = 0

    If limit > 0 Then
        explode = Split(string,delimiter,limit)
    Else
        explode = Split(string,delimiter)
    End If

End Function