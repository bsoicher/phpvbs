'=======================================================================
' ������𔽕�����
'=======================================================================
'�y�����z
'  input        = string    �J��Ԃ�������B
'  multiplier   = int       input ���J��Ԃ��񐔁Bmultiplier �� 0 �ȏ�łȂ���΂Ȃ�܂���B multiplier �� 0 �ɐݒ肳�ꂽ�ꍇ�A���̊֐��͋󕶎���Ԃ��܂��B
'�y�߂�l�z
'  �J��Ԃ����������Ԃ��܂��B
'�y�����z
'  �Einput  �� multiplier  ����J��Ԃ����������Ԃ��܂��B
'=======================================================================
Function str_repeat(input, multiplier)
    If multiplier < 0 Then Exit Function
    str_repeat = String(multiplier,input)
End Function