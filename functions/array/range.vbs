'=======================================================================
'����͈͂̐�����L����z����쐬����
'=======================================================================
'�y�����z
'  low  = mixed �����l�B
'  high = mixed ����l�B
'  step = mixed step  ���w�肳��Ă���ꍇ�A����� �v�f���̑������ƂȂ�܂��Bstep  �͐��̐��łȂ���΂Ȃ�܂���B�f�t�H���g�� 1 �ł��B
'�y�߂�l�z
'  low  ���� high  �܂ł̐����̔z���Ԃ��܂��B low > high �̏ꍇ�A���Ԃ� high ���� low �ƂȂ�܂��B
'�y�����z
'  �E����͈͂̐�����L����z����쐬���܂��B
'=======================================================================
Function range(low,high,step)

    Dim matrix
    Dim inival, endval, plus
    Dim walker : If len(step) > 0 Then walker = step Else walker = 1
    Dim chars : chars = false

    If isNumeric(low) and isNumeric(high) Then
        inival = low
        endval = high
    ElseIf Not isNumeric(low) and Not isNumeric(high) Then
        chars  = true
        inival = Asc(low)
        endval = Asc(high)
    Else
        inival = [?](isNumeric(low),low,0)
        endval = [?](isNumeric(high),high,0)
    End If

    plus = true
    If inival > endval Then plus = false

    If plus Then
        Do While inival <= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival + walker
        Loop
    Else
        Do While inival >= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival - walker
        Loop
    End If

    range = matrix

End Function
