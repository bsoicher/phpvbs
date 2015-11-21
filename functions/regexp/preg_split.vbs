'=======================================================================
'���K�\���ŕ�����𕪊�����
'=======================================================================
'�y�����z
'  pattern      = string ��������p�^�[����\��������B
'  subject      = string ���͕�����B
'  limit        = int    ������w�肵���ꍇ�A�ő� limit  �̕��������񂪕Ԃ���܂��B
'  flags        = int    flags  �́A�t���O��g�ݍ��킹�����̂Ƃ��� �i�r�b�g�a���Z�q�b�őg�ݍ��킹��j���Ƃ��\�ł��B
'�y�߂�l�z
'  pattern  �Ƀ}�b�`�������E�ŕ������� subject  �̕���������̔z���Ԃ��܂��B
'�y�����z
'  �E�w�肵����������A���K�\���ŕ������܂��B
'=======================================================================
Const PREG_SPLIT_NO_EMPTY       = 1
Const PREG_SPLIT_DELIM_CAPTURE  = 2
Const PREG_SPLIT_OFFSET_CAPTURE = 4

Function preg_split(pattern, subject,limit,flags)

    If is_empty(limit) Then limit = 0

    Dim key,matches,tmp_sp,tmp_str
    Dim cnt,counter,strMid,pointer : pointer = 1 : counter = 0
    Dim strRegExp,intPoint,strPoint

    cnt = preg_match_all(pattern,subject, matches, PREG_OFFSET_CAPTURE, "")
    If cnt > 0 Then
        For key = 0 to uBound(matches(0))

            counter = counter + 1
            If limit > 0 Then
                If counter >= limit Then Exit For
            End If

            intPoint  = matches(0)(key)(1)
            strPoint  = matches(0)(key)(0)
            strRegExp = Mid(subject, pointer,intPoint-pointer+1)

            Select Case flags
            Case PREG_SPLIT_NO_EMPTY
                if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
            Case PREG_SPLIT_DELIM_CAPTURE
                [] tmp_sp , strRegExp
                [] tmp_sp , matches(1)(key)(0)
            Case PREG_SPLIT_OFFSET_CAPTURE
                [] tmp_sp , array(strRegExp,pointer-1)
            Case Else
                [] tmp_sp , strRegExp
            End Select
            pointer = intPoint + 1 + len(strPoint)
        Next

        strRegExp = Mid(subject, pointer)
        Select Case flags
        Case PREG_SPLIT_NO_EMPTY
            if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
        Case PREG_SPLIT_DELIM_CAPTURE
            [] tmp_sp , strRegExp
        Case PREG_SPLIT_OFFSET_CAPTURE
            [] tmp_sp , array(strRegExp,pointer-1)
        Case Else
            [] tmp_sp , strRegExp
        End Select
    End If

    preg_split = tmp_sp

End Function