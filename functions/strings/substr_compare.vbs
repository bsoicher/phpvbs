'=======================================================================
' �w�肵���ʒu����w�肵�������� 2 �̕�����ɂ��āA�o�C�i���Ή��Ŕ�r����
'=======================================================================
'�y�����z
'  main_str             = string    �ŏ��̕�����B
'  str                  = string    ���̕�����B
'  offset               = int       ��r���J�n����ʒu�B ���̒l���w�肵���ꍇ�́A������̍Ōォ�琔���܂��B
'  length               = int       ��r���钷���B
'  case_insensitivity   = bool      case_insensitivity  �� TRUE �̏ꍇ�A �啶������������ʂ����ɔ�r���܂��B
'�y�߂�l�z
'  main_str  �� offset  �ȍ~�� str  ��菬�����ꍇ�ɕ��̐��A str  ���傫���ꍇ�ɐ��̐��A �������ꍇ�� 0 ��Ԃ��܂��B
'�y�����z
'  �E substr_compare() �́Amain_str  �� offset  �����ڈȍ~�̍ő� length  �������Astr  �Ɣ�r���܂��B
'=======================================================================
Function substr_compare(main_str,str,offset,length, case_insensitivity)

    If len(offset) > 0 Then
        If offset > 0 Then
            main_str = Mid(main_str,offset)
        Else
            main_str = Mid(main_str,len(main_str) + offset + 1)
        End If
    End If

    If len(length) > 0 Then
        main_str = Left(main_str,length)
        str = Left(str,length)
    End If
    var_dump main_str
    var_dump str
    If case_insensitivity = true Then
        substr_compare = strcasecmp(main_str,str)
    Else
        substr_compare = strcmp(main_str,str)
    End If

End Function