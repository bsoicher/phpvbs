'=======================================================================
' �啶������������ʂ��Ȃ� strstr()
'=======================================================================
'�y�����z
'  haystack     = string    �������s��������B
'  needle       = string    needle �́A �ЂƂ܂��͕����̕����ł��邱�Ƃɒ��ӂ��܂��傤�Bneedle ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  before_needle= string    TRUE �ɂ���� (�f�t�H���g�� FALSE �ł�)�Astristr()  �̕Ԃ�l�́Ahaystack  �̒��ōŏ��� needle  ���������ӏ����O�̕����ƂȂ�܂��B
'�y�߂�l�z
'  �}�b�`���������������Ԃ��܂��Bneedle  ��������Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Ehaystack  �ɂ����� needle  ���ŏ��Ɍ��������ʒu����Ō�܂ł�Ԃ��܂��B
'  �Eneedle  ����� haystack  �͑啶������������ʂ����ɕ]������܂��B
'=======================================================================
Function stristr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbTextCompare)

    If pos <= 0 Then
        stristr = false
    Else
        If before_needle Then
            stristr = Mid(haystack,1,pos-1)
        Else
            stristr = Mid(haystack,pos)
        End If
    End If

End Function