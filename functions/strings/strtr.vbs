'=======================================================================
' ����̕�����ϊ�����
'=======================================================================
'�y�����z
'  str      = string    �ϊ����镶����B
'  from     = string    strTo  �ɕϊ�����镶����B
'  strTo    = int       from  ��u�����镶����B
'�y�߂�l�z
'  ���̊֐��� str  �𑖍����A from  �Ɋ܂܂�镶����������ƁA���̂��ׂĂ� strTo  �̒��őΉ����镶���ɒu�������A ���̌��ʂ�Ԃ��܂��B
'�y�����z
'  �E ���̊֐��� str  �𑖍����A from  �Ɋ܂܂�镶����������ƁA���̂��ׂĂ� to  �̒��őΉ����镶���ɒu�������A ���̌��ʂ�Ԃ��܂��B
'  �E from �� to �̒������قȂ�ꍇ�A�������̗]���ȕ����͖�������܂��B 
'=======================================================================
Function strtr(ByVal str, from, strTo)

    If isObject(from) Then
        Dim key
        For Each key In from
            str = Replace(str,key,from(key))
        Next

    Else

        Dim len1 : len1 = len(from)
        Dim len2 : len2 = len(strTo)

        If len1 > len2 Then
            from = Left(from,len2)
        ElseIf len2 > len1 Then
            strTo = Left(strTo,len1)
        End If

        str = Replace(str,from,strTo)
    End if

    strtr = str
End Function