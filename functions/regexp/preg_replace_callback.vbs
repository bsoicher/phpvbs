'=======================================================================
'���K�\����������ђu�����s��
'=======================================================================
'�y�����z
'  pattern      = mixed �������s���p�^�[���B������������͔z��Ƃ��邱�Ƃ��ł��܂��B
'  callback     = mixed ���̃R�[���o�b�N�֐��́A�����Ώە�����Ń}�b�`�����v�f�̔z�񂪎w�肳��� �R�[������܂��B���̃R�[���o�b�N�֐��́A�u����̕������Ԃ��K�v������܂��B
'  subject      = mixed �����E�u���ΏۂƂȂ镶����������͕�����̔z��
'  limit        = int   subject  ������ɂ����āA�e�p�^�[���ɂ�� �u�����s���ő�񐔁B�f�t�H���g�� -1 (��������)�B
'  cnt          = int   ���̈������w�肳���ƁA�u���񐔂��n����܂��B
'�y�߂�l�z
'  subject  �������z��̏ꍇ�͔z����A ���̑��̏ꍇ�͕������Ԃ��܂��B
'  �p�^�[�����}�b�`�����ꍇ�A�k�u�����s��ꂽ�l�V���� subject  ��Ԃ��܂��B
'  �}�b�`���Ȃ������ꍇ�Asubject  �����̂܂ܕԂ��܂��B
'�y�����z
'  �Esubject  �Ɋւ��� pattern  ��p���Č������s���A callback  �ɒu�����܂��B
'=======================================================================
Function preg_replace_callback(pattern,callback,ByVal subject,limit,ByRef cnt)

    Dim key,counter
    cnt = 0
    If len(limit) = 0 Then limit = 0

    If isArray(subject) Then
        For key = 0 to uBound(subject)
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    ElseIf isObject(subject) Then
        For Each key In subject
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    Else

        If isArray(pattern) Then
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), callback, _
                                   subject,limit,cnt)
                Next

        Else

            Dim matchAll,strCallback
            If is_empty(limit) Then
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each key In matchAll(0)
                        execute("strCallback = " & callback & "(key)")
                        subject = Replace(subject,key,strCallback)
                    Next
                End If

            Else
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each counter In matchAll(0)
                        cnt = cnt + 1
                        If cnt > limit Then Exit For
                        execute("strCallback = " & callback & "(counter)")
                        subject = Replace(subject,counter,strCallback)
                    Next
                End If
            End If

        End If
    End If

    preg_replace_callback = subject

End Function