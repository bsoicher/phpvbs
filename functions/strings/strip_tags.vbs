'=======================================================================
' �����񂩂� HTML�^�O����菜��
'=======================================================================
'�y�����z
'  str              = string ���͕�����B
'  allowable_tags   = string �I�v�V������2�Ԗڂ̈����ɂ��A��菜���Ȃ��^�O���w��ł��܂��B
'�y�߂�l�z
'  �^�O�����������������Ԃ��܂��B
'�y�����z
'  �E�w�肵�������� ( str ) ����S�Ă� HTML�^�O����菜���܂��B
'=======================================================================
Function strip_tags( str )

    Dim objRegExp
    Dim plane

    plane = Trim( str & "" )

    If Len( plane ) > 0 Then

        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        objRegExp.Pattern= "</?[^>]+>"
        plane = objRegExp.Replace(str, "")
        Set objRegExp = Nothing

    End If

    strip_tags = plane

End Function