'=======================================================================
' ������̈ꕔ����Ԃ�
'=======================================================================
'�y�����z
'  str          = string    ���͕�����B
'  start        = string     start  �����̏ꍇ�A�Ԃ���镶����́A string  �� 0 ���琔���� start �Ԗڂ���n�܂镶����ƂȂ�܂��B �Ⴆ�΁A������'abcdef'�ɂ����Ĉʒu 0�ɂ��镶���́A'a'�ł���A �ʒu2�ɂ�'c'������܂��Bstart �����̏ꍇ�A�Ԃ���镶����́A string �̌�납�琔���� start �Ԗڂ���n�܂镶����ƂȂ�܂��B 
'  intLength    = string    ���͕�����B
'�y�߂�l�z
'  ������̈ꕔ��Ԃ��܂��B
'�y�����z
'  �E������ str  �́Astart  �Ŏw�肳�ꂽ�ʒu���� length  �o�C�g���̕������Ԃ��܂��B
'=======================================================================
Function substr(ByVal str,ByVal start,ByVal intLength)

	intStart = start
    If len(intStart) = 0 Then intStart = 0
    if intStart = 0 And Len(intLength) < 1 Then
    	substr = str
    	Exit Function
    End If
    
    If len(intLength) = 0 Then intLength = abs(start)
    
    If intStart < 0 Then
        intStart = len(str)+1 + intStart
    Else
    	intStart = intStart + 1
    End If

    Dim tmp
    If intStart <> 0 Then
	    tmp = Mid(str,intStart)
	Else 
		tmp = str
	End If

    Dim intLen
    intLen = len(tmp)


    If len(intLength) > 0 Then

        If intLen >= abs(intLength) Then
            If intLength > 0 Then

        		tmp = Left(tmp,intLength)

            Else
                tmp = Left(tmp,len(tmp) + intLength)
            End If
        Else
            tmp = False
        End If
    End If

    substr = tmp

End Function