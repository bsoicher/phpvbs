'=======================================================================
'�z�񂩂��ȏ�̗v�f�������_���Ɏ擾����
'=======================================================================
'�y�����z
'  mAry     = array  ���͂̔z��B
'  num_req  = int    �擾����G���g���̐����w�肵�܂��B �w�肳��Ȃ��ꍇ�́A�f�t�H���g�� 1 �ɂȂ�܂��B
'�y�߂�l�z
'  �G���g����������擾����ꍇ�A array_rand() �̓����_���ȃG���g���̃L�[��Ԃ��܂��B
'  ���̑��̏ꍇ�́A�����_���ȃG���g���̃L�[�̔z���Ԃ��܂��B 
'  ����ɂ��A�����_���ȃL�[���擾���A �z�񂩂�l���擾���邱�Ƃ��\�ɂȂ�܂��B
'�y�����z
'  �E�z�񂩂��ȏ�̃����_���ȃG���g�����擾���悤�Ƃ���ꍇ�ɗL�p�ł��B
'=======================================================================
Function array_rand(mAry, ByVal num_req)

    If Not isArray( mAry ) Then Exit Function
    If Not isNumeric( num_req ) Then num_req = 1

    Dim rand,i,intCounter,aryCounter,indexes

    intCounter = uBound(mAry)
    aryCounter = num_req -1

    If intCounter < aryCounter Then Exit Function

    Randomize

    ReDim indexes( aryCounter )
    For i = 0 to aryCounter
        Do While true
            rand = Round( Rnd * uBound(mAry) )
            If Not in_array(rand, indexes,true) Then
                indexes(i) = rand
                Exit Do
            End If
        Loop
    Next

    If num_req = 1 Then
        array_rand = indexes(0)
    Else
        array_rand = indexes
    End If

End Function