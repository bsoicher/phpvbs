'=======================================================================
' ���ꕶ���� HTML �G���e�B�e�B�ɕϊ�����
'=======================================================================
'�y�����z
'  str  = string    �ϊ�����镶����B
'�y�߂�l�z
'  �ϊ���̕������Ԃ��܂��B
'�y�����z
'  �E�����̒��ɂ� HTML �ɂ����ē���ȈӖ��������̂�����A �����̖{���̒l��\����������� HTML �̕\���`���ɕϊ����Ă��Ȃ���΂Ȃ�܂���B
'  �E���̊֐��́A�����̕ϊ����s�������ʂ̕������Ԃ��܂��B 
'=======================================================================
Const ENT_NOQUOTES = 0
Const ENT_COMPAT   = 2
Const ENT_QUOTES   = 3
Function htmlSpecialChars(ByVal str)

    if len( str ) > 0 then
        str = Server.HTMLEncode(str)
        str = Replace(str,"'","&#039;")
    end if
    htmlspecialchars = str

End Function