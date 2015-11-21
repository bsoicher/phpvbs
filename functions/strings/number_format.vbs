'=======================================================================
' �������ʖ��ɃO���[�v�����ăt�H�[�}�b�g����
'=======================================================================
'�y�����z
'  number           = float     �t�H�[�}�b�g���鐔�l�B
'  decimals         = int       �����_�ȉ��̌����B
'  dec_point        = string    �����_��\����؂蕶���B
'  thousands_sep    = string    ��ʖ��̋�؂蕶���Bthousands_sep �͍ŏ��̕����������g�p����܂��B �Ⴆ�΁A������ 1000 �ɑ΂��� thousands_sep �Ƃ��� bar ���g�p�����ꍇ�Anumber_format() �� 1b000 ��Ԃ��܂��B
'�y�߂�l�z
'  �ύX��̕������Ԃ��܂��B
'�y�����z
'  �E�p�����[�^�� 1 �����n���ꂽ�ꍇ�A number  �͐�ʖ��ɃJ���} (",") ���ǉ�����A �����Ȃ��Ńt�H�[�}�b�g����܂��B
'  �E�p�����[�^�� 2 �n���ꂽ�ꍇ�Anumber �� decimals ���̏����̑O�Ƀh�b�g (".") �A ��ʖ��ɃJ���} (",") ���ǉ�����ăt�H�[�}�b�g����܂��B
'  �E�p�����[�^�� 4 �S�ēn���ꂽ�ꍇ�Anumber �̓h�b�g (".") �̑���� dec_point �� decimals ���̏����̑O�ɁA��ʖ��ɃJ���} (",") �̑���� thousands_sep ���ǉ�����ăt�H�[�}�b�g����܂��B 
'=======================================================================
Function number_format( number, decimals, dec_point, thousands_sep )

    Dim n,c,d,t,i,s
    n = number
    c = [?]( isNumeric(decimals),decimals,2 )
    c = abs( c )

    d = [?]( len(dec_point) = 0,",",dec_point)
    t = [?]( len(thousands_sep) = 0, ".", left(thousands_sep,1) )

    n = FormatNumber (n, c,true,false,true)
    n = Replace(n,",",d)
    n = Replace(n,".",t)

    number_format = n

End Function
