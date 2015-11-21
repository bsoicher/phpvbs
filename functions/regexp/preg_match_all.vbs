'=======================================================================
'�J��Ԃ����K�\���������s��
'=======================================================================
'�y�����z
'  pattern  = string    ��������p�^�[����\��������B
'  subject  = string    ���͕�����B
'  matches  = array     matches  ���w�肵���ꍇ�A�������ʂ��������܂��Bmatches(0) �ɂ̓p�^�[���S�̂Ƀ}�b�`�����e�L�X�g���������A matches(1)�ɂ� 1 �Ԗڂ̂̃L���v�`���p�T�u�p�^�[���Ƀ}�b�`���� �����񂪑������A�Ƃ������悤�ɂȂ�܂��B
'  flags    = int       �߂�l�̌`�����w��
'  offset   = int       �ʏ�A�����͑Ώە�����̐擪����J�n����܂��B �I�v�V�����̃p�����[�^ offset  ���g�p���� �����̊J�n�ʒu���w�肷�邱�Ƃ��\�ł��B
'�y�߂�l�z
'  �p�^�[�����}�b�`����������Ԃ��܂��i�[���ƂȂ�\��������܂��j�B 
'  �܂��́A�G���[�����������ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E subject  ���������A pattern  �Ɏw�肵�����K�\���Ƀ}�b�`���� ���ׂĂ̕�������Aflags  �Ŏw�肵�� ���ԂŁAmatches  �ɑ�����܂��B
'  �E ���K�\���Ƀ}�b�`����ƁA���̃}�b�`����������̌ォ�� ���������s����܂��B 
'=======================================================================
Function preg_match_all(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim regEx,matchall,matchone
    Dim cnt,counter : counter = 0
    Dim helper

    preg_match_all = false
    If vartype(matches) <> 0 Then Exit Function
    If len(flags) = 0 Then flags = PREG_PATTERN_ORDER

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set regEx = new RegExp
    With regEx
        .IgnoreCase = helper.withIgnoreCase
        .Global     = True
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With
    set helper = Nothing

    If len(offset) > 0 Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    Set matchall = regEx.execute(subject)
    Set regEx = Nothing
    If matchall.count = 0 Then exit Function

    If flags = PREG_PATTERN_ORDER Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            matches(0)(counter) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(cnt)(counter) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    Elseif flags = PREG_SET_ORDER Then

        ReDim matches(matchall.count-1)

        counter = 0
        For Each matchone In matchall
            toReDim matches(counter),(matchone.SubMatches.count)
            matches(counter)(0) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(counter)(cnt) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    ElseIf PREG_OFFSET_CAPTURE Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            toReDim matches(0)(counter),1
            matches(0)(counter)(0) = matchone.value
            matches(0)(counter)(1) = matchone.FirstIndex
            For cnt = 1 to matchone.SubMatches.count
                toReDim matches(cnt)(counter),1
                matches(cnt)(counter)(0) = matchone.SubMatches(cnt-1)
                matches(cnt)(counter)(1) = InStr( matchone.value, matchone.SubMatches(cnt-1) ) -1
            Next
            counter = counter + 1
        Next

    End If

    preg_match_all = matchall.Count

End Function