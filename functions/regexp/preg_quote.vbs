'=======================================================================
'���K�\���������N�I�[�g����
'=======================================================================
'�y�����z
'  str          = string    ���͕�����B
'  delimiter    = string    �I�v�V������ delimiter  ���w�肷��ƁA �����Ŏw�肵���������G�X�P�[�v����܂��B����́APCRE �֐����g�p���� �f���~�^���G�X�P�[�v����ꍇ�ɕ֗��ł��B'/' ���f���~�^�Ƃ��Ă� �ł���ʓI�Ɏg�p����Ă��܂��B
'�y�߂�l�z
'  �N�H�[�g���ꂽ�������Ԃ��܂��B
'�y�����z
'  �E preg_quote() �́Astr  �������Ƃ��A���K�\���\���̓��ꕶ���̑O�Ƀo�b�N�X���b�V����}�����܂��B
'  �E ���̊֐��́A���s���ɐ�������镶������p�^�[���Ƃ��ă}�b�`���O���s���K�v������A ���̕�����ɂ͐��K�\���̓��ꕶ�����܂܂�Ă��邩���m��Ȃ��ꍇ�ɗL�p�ł��B
'  �E ���K�\���̓��ꕶ���́A���̂��̂ł��B . \ + * ? [ ^ ] $ ( ) { } = ! < > | : 
'=======================================================================
Function preg_quote(ByVal str,delimiter)

    Dim pattern : pattern = array("\",".","+","*","?","[","^","]","$","(",")","{","}","=","!","<",">","|",":")
    If len(delimiter) > 0 Then [] pattern , delimiter

    Dim key
    For key = 0 to uBound(pattern)
        str = Replace(str,pattern(key),"\" & pattern(key))
    Next
    preg_quote = str

End Function