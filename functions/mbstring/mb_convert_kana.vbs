'BASP21���g�p���Ȃ��ꍇ
'http://www.ac.cyberhome.ne.jp/~mattn/AcrobatASP/1.html
'http://www.ac.cyberhome.ne.jp/~mattn/AcrobatASP/StrConv.inc
'=======================================================================
' �J�i��("�S�p����"�A"���p����"����)�ϊ�����
'=======================================================================
'�y�����z
'  str      = string  ������
'  option   = mixed   �ϊ��I�v�V����
'�y�߂�l�z
'  �ϊ����ꂽ�������Ԃ��܂��B
'�y�����z
'  �E������ str  �Ɋւ��āu���p�v-�u�S�p�v�ϊ����s���܂��B
'=======================================================================
Function mb_convert_kana( str , option_name )

    Dim outstr
    Dim J
    Dim tmp
    Dim bobj

    If Len( str ) = 0 Then Exit Function

    Set bobj = Server.CreateObject("basp21")

    If Len( option_name ) = 0 Then
        option_name = "K"
    End If

    outstr = str

    For J = 1 To Len( option_name )

        tmp = mid(option_name,J,1)

        If tmp = "r" Then
            '�S�p���� (2 �o�C�g) �𔼊p���� (1 �o�C�g) �ɕϊ�
            outstr = bobj.StrConv(outstr,8)

        ElseIf tmp = "R" Then
            '���p���� (1 �o�C�g) ��S�p���� (2 �o�C�g) �ɕϊ�
            outstr = bobj.StrConv(outstr,4)

        ElseIf tmp = "c" Then
            '�J�^�J�i���Ђ炪�Ȃɕϊ�
            outstr = bobj.StrConv(outstr,32)

        ElseIf tmp = "C" Then
            '�Ђ炪�Ȃ��J�^�J�i�ɕϊ�
            outstr = bobj.StrConv(outstr,16)

        ElseIf tmp = "K" or tmp = "V" THen
            '������̒��̔��p�J�i��S�p�J�i�ɕϊ����܂��B�����ɂ��Ή��B
            outstr = bobj.HAN2ZEN(outstr)

        End If
    Next

    mb_convert_kana = outstr

End Function