'=======================================================================
' �啶������������ʂ����ɕ����񂪍ŏ��Ɍ����ʒu��T��
'=======================================================================
'�y�����z
'  haystack = string    �������s��������B
'  needle   = string    needle �́A �ЂƂ܂��͕����̕����ł��邱�Ƃɒ��ӂ��܂��傤�Bneedle ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  offset   = string    �I�v�V�����̃p�����[�^ offset  �ɂ��A�������J�n���� haystack  �̈ʒu���w�肷�邱�Ƃ��ł��܂��B ���̏ꍇ�ł��Ԃ����ʒu�́A haystack  �̐擪����̈ʒu�̂܂܂ƂȂ�܂��B
'�y�߂�l�z
'  needle  ���݂���Ȃ��ꍇ�A stripos() �� boolean FALSE  ��Ԃ��܂��B
'�y�����z
'  �E ������ haystack  �̒��� needle  ���ŏ��Ɍ����ʒu�𐔎��ŕԂ��܂��B
'  �E strpos() �ƈقȂ�Astripos() �͑啶������������ʂ��܂���B 
'=======================================================================
Function stripos( haystack, needle, offset)

    Dim i
    stripos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbTextCompare)

    If i > 0 Then
        stripos = i
    End If

End Function