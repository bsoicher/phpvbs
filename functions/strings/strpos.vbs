'=======================================================================
' �����񂪍ŏ��Ɍ����ꏊ��������
'=======================================================================
'�y�����z
'  haystack = string    �������s��������B
'  needle   = string   needle  ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  offset   = string    �I�v�V�����̃p�����[�^ offset  �ɂ��A�������J�n���� haystack  �̈ʒu���w�肷�邱�Ƃ��ł��܂��B ���̏ꍇ�ł��Ԃ����ʒu�́A haystack  �̐擪����̈ʒu�̂܂܂ƂȂ�܂��B
'�y�߂�l�z
'  �ʒu��\�������l��Ԃ��܂��B needle  ��������Ȃ��ꍇ�A strpos() �� boolean FALSE ��Ԃ��܂��B
'�y�����z
'  �E ������ haystack  �̒��ŁA needle  ���ŏ��Ɍ��ꂽ�ʒu�𐔎��ŕԂ��܂��B
'  �E PHP 5 �ȑO�� strrpos() �Ƃ͈قȂ�A���̊֐��� needle  �p�����[�^�Ƃ��ĕ�����S�̂��Ƃ�A ���̕�����S�̂������ΏۂƂȂ�܂��B
'=======================================================================
Function strpos( haystack, needle, offset)

    Dim i
    strpos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbBinaryCompare)

    If i > 0 Then
        strpos = i
    End If

End Function