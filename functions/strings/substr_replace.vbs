'=======================================================================
' ������̈ꕔ��u������
'=======================================================================
'�y�����z
'  str          = string    ���͕�����B
'  replacement  = string    �u�����镶����B
'  start        = string    start �����̏ꍇ�A�u���� string �� start �Ԗڂ̕�������n�܂�܂��Bstart �����̏ꍇ�A�u���� string �̏I�[���� start �Ԗڂ̕�������n�܂�܂��B
'  length       = string    ���̒l���w�肵���ꍇ�A string �@�̒u������镔���̒�����\���܂��B ���̏ꍇ�A�u�����~����ʒu�� string  �̏I�[���牽�����ڂł��邩��\���܂��B���̃p�����[�^���ȗ����ꂽ�ꍇ�A �f�t�H���g�l�� strlen(string )�A���Ȃ킿�A string  �̏I�[�܂Œu�����邱�ƂɂȂ�܂��B ���R�A���� length  ���[����������A ���̊֐��� string  �̍ŏ����� start  �̈ʒu�� replacement  ��}������Ƃ������ƂɂȂ�܂��B
'�y�߂�l�z
'  ���ʂ̕������Ԃ��܂��B�����Astring  ���z��̏ꍇ�A�z�񂪕Ԃ���܂��B
'�y�����z
'  �Esubstr_replace()�́A������ string �� start  ����� (�I�v�V������) length  �p�����[�^�ŋ�؂�ꂽ������ replacement  �Ŏw�肵��������ɒu�����܂��B
'=======================================================================
Function substr_replace(str, replacement, start, length)

    Dim key

    If isArray(str) Then
        For key = 0 to uBound(str)
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    ElseIf isObject(str) Then
        For Each key In str
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    End If

    Dim result

    If start < 0 Then
        start = len(str) + start
    End If

    If start <> 0 Then
        result = Left(str,start)
    End If

    result = result & replacement

    If len(length) > 0 Then
        If length > 0 Then
            result = result & Mid(str,start + length)
        Else
            result = result & Right(str,abs(length))
        End If

    End If

    substr_replace = result
End Function