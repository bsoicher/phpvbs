'=======================================================================
'���K�\����������ђu�����s��
'=======================================================================
'�y�����z
'  pattern      = mixed �������s���p�^�[���B������������͔z��Ƃ��邱�Ƃ��ł��܂��B
'  replacement  = mixed �u�����s��������������͕�����̔z��B
'  subject      = mixed �����E�u���ΏۂƂȂ镶����������͕�����̔z��
'  limit        = int   subject  ������ɂ����āA�e�p�^�[���ɂ�� �u�����s���ő�񐔁B�f�t�H���g�� -1 (��������)�B
'  cnt          = int   ���̈������w�肳���ƁA�u���񐔂��n����܂��B
'�y�߂�l�z
'  subject  �������z��̏ꍇ�͔z����A ���̑��̏ꍇ�͕������Ԃ��܂��B
'  �p�^�[�����}�b�`�����ꍇ�A�k�u�����s��ꂽ�l�V���� subject  ��Ԃ��܂��B
'  �}�b�`���Ȃ������ꍇ�Asubject  �����̂܂ܕԂ��܂��B
'�y�����z
'  �Esubject  �Ɋւ��� pattern  ��p���Č������s���A replacement  �ɒu�����܂��B
'=======================================================================
Function preg_replace(pattern,replacement,ByVal subject,limit,ByRef cnt)

    Dim key
    cnt = 0

    If isArray(subject) Then
        For key = 0 to uBound(subject)
            subject(key) = preg_replace( pattern, replacement, subject(key),limit,cnt)
        Next
    ElseIf isObject(subject) Then
        For Each key In subject
            subject(key) = preg_replace( pattern, replacement, subject(key),limit,cnt)
        Next
    Else

        If isArray(pattern) Then
            If not isArray(replacement) Then
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), replacement, _
                                   subject,limit,cnt)
                Next
            ElseIf isArray(replacement) Then

                If uBound(pattern) <> uBound(replacement) Then
                    ReDim Preserve replacement(uBound(pattern))
                End If
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), replacement(key), _
                                   subject,limit,cnt)
                Next
            End If

        Else

            Dim strRetValue, RegEx,helper

            set helper = new RegExp_Helper
            helper.parseOption(pattern)

            Set RegEx = New RegExp
            With RegEx
                .IgnoreCase = helper.withIgnoreCase
                .Global     = [?](limit,false,true)
                .pattern    = helper.withPattern
                .MultiLine  = helper.withMultiLine
            End With
            set helper = Nothing

            If RegEx.Global Then
                If  len(subject) > 0 Then _
                subject = RegEx.Replace(subject, replacement)
            Else
                For key = 1 to limit
                    subject = RegEx.Replace(subject, replacement)
                    cnt = cnt + 1
                Next
            End If

            set RegEx = Nothing

        End If
    End If

    preg_replace = subject

End Function