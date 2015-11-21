'=======================================================================
' ��������Œ蒷�̑��̕�����Ŗ��߂�
'=======================================================================
'�y�����z
'  input        = string    ���͕�����B
'  pad_length   = int       pad_length  �̒l�����A �܂��͓��͕�����̒��������Z���ꍇ�A���߂鑀��͍s���܂���B
'  pad_string   = string    �K�v�Ƃ���閄�߂镶������ pad_string  �̒����ŋϓ��ɕ����ł��Ȃ��ꍇ�Apad_string  �͐؂�̂Ă��܂��B 
'  pad_type     = int       �I�v�V�����̈��� pad_type  �ɂ́A STR_PAD_RIGHT, STR_PAD_LEFT, STR_PAD_BOTH  ���w��\�ł��B pad_type ���w�肳��Ȃ��ꍇ�A STR_PAD_RIGHT  �����肵�܂��B
'�y�߂�l�z
'  �t�H�[�}�b�g������ format  �Ɋ�Â��������ꂽ�������Ԃ��܂��B
'�y�����z
'  �E���̊֐��͕����� input  �̍��A�E�܂��͗������w�肵�������Ŗ��߂܂��B�I�v�V�����̈��� pad_string  ���w�肳��Ă��Ȃ��ꍇ�́A input  �͋󔒂Ŗ��߂��A����ȊO�̏ꍇ�́A pad_string  ����̕����Ő����܂Ŗ��߂��܂��B
'=======================================================================
Const STR_PAD_LEFT  = 0
Const STR_PAD_RIGHT = 1
Const STR_PAD_BOTH  = 2
Function str_pad(byVal input, pad_length, pad_string, pad_type)

    Dim half : half = ""
    Dim pad_to_go

    If pad_type <> STR_PAD_LEFT and pad_type <> STR_PAD_RIGHT and pad_type <> STR_PAD_BOTH Then
        pad_type = STR_PAD_RIGHT
    End If

    If len(pad_string) = 0 Then pad_string = " "

    pad_to_go = pad_length - len( input )
    If pad_to_go > 0 Then
        If pad_type = STR_PAD_LEFT Then
            input = str_pad_helper(pad_string, pad_to_go) & input
        ElseIf pad_type = STR_PAD_RIGHT Then
            input = input & str_pad_helper(pad_string, pad_to_go)
        ElseIf pad_type = STR_PAD_BOTH Then
            half = str_pad_helper(pad_string,intval(pad_to_go/2))
            input = half & input & half
            input = Left(input,pad_length)
        End If
    End if

    str_pad = input

End Function

'***************************
Function str_pad_helper(s, intlen)

    Dim collect : collect = ""
    Dim i

    Do Until len( collect ) > intlen
        collect = collect & s
    Loop

    collect = Left(collect,intlen)

    str_pad_helper = collect

End Function