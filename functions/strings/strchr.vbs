'=======================================================================
' strstr() �̃G�C���A�X
'=======================================================================
'�y�����z
'  haystack     = string    ���͕�����B
'  needle       = string    needle  ��������łȂ��ꍇ�́A ����𐮐��ɕϊ����A���̔ԍ��ɑΉ����镶���Ƃ��Ĉ����܂��B
'  before_needle= string    TRUE �ɂ���� (�f�t�H���g�� FALSE �ł�)�Astrstr()  �̕Ԃ�l�́Ahaystack  �̒��ōŏ��� needle  ���������ӏ����O�̕����ƂȂ�܂��B
'�y�����z
'  �E���̊֐��͎��̊֐��̃G�C���A�X�ł��B strstr().
'=======================================================================
Function strchr( haystack, needle, before_needle )
    strchr = strstr( haystack, needle, before_needle )
End Function