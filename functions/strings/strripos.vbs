'=======================================================================
' �����񒆂ŁA�����(�啶������������ʂ��Ȃ�)�����񂪍Ō�Ɍ��ꂽ�ʒu��T��
'=======================================================================
'�y�����z
'  haystack = string    �������s��������B
'  needle   = string   needle  ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  offset   = string    �I�v�V�����̃p�����[�^ offset  �ɂ��A�������J�n���� haystack  �̈ʒu���w�肷�邱�Ƃ��ł��܂��B ���̏ꍇ�ł��Ԃ����ʒu�́A haystack  �̐擪����̈ʒu�̂܂܂ƂȂ�܂��B
'�y�߂�l�z
'  needle  ���Ō�Ɍ��ꂽ�ʒu��Ԃ��܂��B
'  needle  ��������Ȃ��ꍇ�AFALSE ��Ԃ��܂��B
'�y�����z
'  �E ������̒��ŁA �啶������������ʂ��Ȃ����镶���񂪍Ō�Ɍ��ꂽ�ʒu��Ԃ��܂��B
'  �E strrpos() �ƈقȂ�Astrripos()  �͑啶������������ʂ��܂���B
'=======================================================================
Function strripos( haystack, needle, offset)

    Dim i
    strripos = false

    If len(offset) = 0 Then
        offset = len( haystack)
    End If

    If len(needle) > 1 Then needle = Left(needle,1)

    i = InStrRev(haystack,needle,offset,vbTextCompare)

    If i > 0 Then
        strripos = i
    End If

End Function