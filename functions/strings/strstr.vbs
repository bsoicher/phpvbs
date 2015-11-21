'=======================================================================
' �����񂪍ŏ��Ɍ����ʒu��������
'=======================================================================
'�y�����z
'  haystack     = string    ���͕�����B
'  needle       = string    needle  ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  before_needle= string    TRUE �ɂ���� (�f�t�H���g�� FALSE �ł�)�Astrstr()  �̕Ԃ�l�́Ahaystack  �̒��ōŏ��� needle  ���������ӏ����O�̕����ƂȂ�܂��B
'�y�߂�l�z
'  �����������Ԃ��܂��B needle  ��������Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Ehaystack  �̒��� needle  ���ŏ��Ɍ����ꏊ���當����̏I���܂ł�Ԃ��܂��B
'=======================================================================
Function strstr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbBinaryCompare)

    If pos <= 0 Then
        strstr = false
    Else
        If before_needle Then
            strstr = Mid(haystack,1,pos-1)
        Else
            strstr = Mid(haystack,pos)
        End If
    End If

End Function