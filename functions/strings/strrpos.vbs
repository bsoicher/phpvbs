'=======================================================================
' �����񒆂ɁA���镶�����Ō�Ɍ����ꏊ��T��
'=======================================================================
'�y�����z
'  haystack = string    �������s��������B
'  needle   = string   needle  ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  offset   = string    �I�v�V�����̃p�����[�^ offset  �ɂ��A�������J�n���� haystack  �̈ʒu���w�肷�邱�Ƃ��ł��܂��B ���̏ꍇ�ł��Ԃ����ʒu�́A haystack  �̐擪����̈ʒu�̂܂܂ƂȂ�܂��B
'�y�߂�l�z
'  needle  ���Ō�Ɍ��ꂽ�ʒu��Ԃ��܂��B
'  needle  ��������Ȃ��ꍇ�AFALSE ��Ԃ��܂��B
'�y�����z
'  �E ������ haystack  �̒��ŁA needle  ���Ō�Ɍ��ꂽ�ʒu�𐔎��ŕԂ��܂��B
'  �E needle  �ɕ����񂪎w�肳�ꂽ�ꍇ�A���̕�����̍ŏ��̕����������g���܂��B
'=======================================================================
Function strrpos( haystack, needle, offset)

    Dim i
    strrpos = false

    If len(offset) = 0 Then
        offset = len( haystack)
    End If

    If len(needle) > 1 Then needle = Left(needle,1)

    i = InStrRev(haystack,needle,offset,vbBinaryCompare)

    If i > 0 Then
        strrpos = i
    End If

End Function