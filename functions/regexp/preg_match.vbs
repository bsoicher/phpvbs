'=======================================================================
'���K�\���ɂ��}�b�`���O���s��
'=======================================================================
'�y�����z
'  pattern  = string    ��������p�^�[����\��������B
'  subject  = string    ���͕�����B
'  matches  = array     matches  ���w�肵���ꍇ�A�������ʂ��������܂��Bmatches(0) �ɂ̓p�^�[���S�̂Ƀ}�b�`�����e�L�X�g���������A matches(1)�ɂ� 1 �Ԗڂ̂̃L���v�`���p�T�u�p�^�[���Ƀ}�b�`���� �����񂪑������A�Ƃ������悤�ɂȂ�܂��B
'  flags    = int       PREG_OFFSET_CAPTURE   ���̃t���O��ݒ肵���ꍇ�A�e�}�b�`�ɑΉ����镶����̃I�t�Z�b�g���Ԃ���܂��B ����ɂ��A�Ԃ�l�͔z��ƂȂ�A�z��̗v�f 0 �̓}�b�`����������A �v�f 1�͑Ώە����񒆂ɂ�����}�b�`����������̃I�t�Z�b�g�l �ƂȂ邱�Ƃɒ��ӂ��Ă��������B
'  offset   = int       �ʏ�A�����͑Ώە�����̐擪����J�n����܂��B �I�v�V�����̃p�����[�^ offset  ���g�p���� �����̊J�n�ʒu���w�肷�邱�Ƃ��\�ł��B
'�y�߂�l�z
'  preg_match() �́Apattern  ���}�b�`�����񐔂�Ԃ��܂��B
'  �܂�A0 ��i�}�b�`�����j�܂��� 1 ��ƂȂ�܂��B
'  ����́A�ŏ��Ƀ}�b�`�������_��preg_match()  �͌������~�߂邽�߂ł��B
'�y�����z
'  �Epattern  �Ŏw�肵�����K�\���ɂ�� subject  ���������܂��B
'=======================================================================
Const PREG_PATTERN_ORDER  = 1
Const PREG_SET_ORDER      = 2
Const PREG_OFFSET_CAPTURE = 256
Function preg_match(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim matchAll,matchone
    Dim cnt,helper

    preg_match = false

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set matchAll = new RegExp
    With matchAll
        .IgnoreCase = helper.withIgnoreCase
        .Global     = false
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With

    set helper = Nothing

    If not is_empty(offset) Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    offset = intval( offset )

    If vartype(matches) <> 8 Then
        Set matchone = matchAll.execute(subject)
        Set matchAll = Nothing
        If matchone.count = 0 Then exit Function

        If flags = PREG_OFFSET_CAPTURE Then

            ReDim matches(1)
            matches(0) = matchone(0).value
            matches(1) = offset + matchone(0).FirstIndex
        Else
            ReDim matches(0)
            matches(0) = matchone(0).value
        End If

        preg_match = true
    Else
        preg_match = matchAll.Test(subject)
        Set matchAll = Nothing
    End If

End Function